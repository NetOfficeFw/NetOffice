using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// Range
	///</summary>
	public class Range_ : COMObject
	{
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(COMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// Get
		/// </summary>
		/// <param name="dataOnly">optional bool DataOnly</param>
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_XML(bool dataOnly)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(dataOnly);
			object returnItem = Invoker.PropertyGet(this, "XML", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// Alias for get_XML
		/// </summary>
		/// <param name="dataOnly">optional bool DataOnly</param>
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public string XML(bool dataOnly)
		{
			return get_XML(dataOnly);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// DispatchInterface Range 
	/// SupportByVersion Word, 9,10,11,12,14,15
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Range : Range_
	{
		#pragma warning disable
		#region Type Information

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Range);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public string Text
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Text", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Text", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range FormattedText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormattedText", paramsArray);
				NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormattedText", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 Start
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Start", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Start", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 End
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "End", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "End", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Font Font
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Font", paramsArray);
				NetOffice.WordApi.Font newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Font.LateBindingApiWrapperType) as NetOffice.WordApi.Font;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Font", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range Duplicate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Duplicate", paramsArray);
				NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdStoryType StoryType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StoryType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdStoryType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Tables Tables
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Tables", paramsArray);
				NetOffice.WordApi.Tables newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Tables.LateBindingApiWrapperType) as NetOffice.WordApi.Tables;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Words Words
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Words", paramsArray);
				NetOffice.WordApi.Words newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Words.LateBindingApiWrapperType) as NetOffice.WordApi.Words;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Sentences Sentences
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Sentences", paramsArray);
				NetOffice.WordApi.Sentences newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Sentences.LateBindingApiWrapperType) as NetOffice.WordApi.Sentences;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Characters Characters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Characters", paramsArray);
				NetOffice.WordApi.Characters newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Characters.LateBindingApiWrapperType) as NetOffice.WordApi.Characters;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Footnotes Footnotes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Footnotes", paramsArray);
				NetOffice.WordApi.Footnotes newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Footnotes.LateBindingApiWrapperType) as NetOffice.WordApi.Footnotes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Endnotes Endnotes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Endnotes", paramsArray);
				NetOffice.WordApi.Endnotes newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Endnotes.LateBindingApiWrapperType) as NetOffice.WordApi.Endnotes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Comments Comments
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Comments", paramsArray);
				NetOffice.WordApi.Comments newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Comments.LateBindingApiWrapperType) as NetOffice.WordApi.Comments;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Cells Cells
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
				NetOffice.WordApi.Cells newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Cells.LateBindingApiWrapperType) as NetOffice.WordApi.Cells;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Sections Sections
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Sections", paramsArray);
				NetOffice.WordApi.Sections newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Sections.LateBindingApiWrapperType) as NetOffice.WordApi.Sections;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Paragraphs Paragraphs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Paragraphs", paramsArray);
				NetOffice.WordApi.Paragraphs newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Paragraphs.LateBindingApiWrapperType) as NetOffice.WordApi.Paragraphs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Borders Borders
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Borders", paramsArray);
				NetOffice.WordApi.Borders newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Borders.LateBindingApiWrapperType) as NetOffice.WordApi.Borders;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Borders", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Shading Shading
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shading", paramsArray);
				NetOffice.WordApi.Shading newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Shading.LateBindingApiWrapperType) as NetOffice.WordApi.Shading;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.TextRetrievalMode TextRetrievalMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextRetrievalMode", paramsArray);
				NetOffice.WordApi.TextRetrievalMode newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.TextRetrievalMode.LateBindingApiWrapperType) as NetOffice.WordApi.TextRetrievalMode;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TextRetrievalMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Fields Fields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Fields", paramsArray);
				NetOffice.WordApi.Fields newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Fields.LateBindingApiWrapperType) as NetOffice.WordApi.Fields;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.FormFields FormFields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormFields", paramsArray);
				NetOffice.WordApi.FormFields newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.FormFields.LateBindingApiWrapperType) as NetOffice.WordApi.FormFields;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Frames Frames
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Frames", paramsArray);
				NetOffice.WordApi.Frames newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Frames.LateBindingApiWrapperType) as NetOffice.WordApi.Frames;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.ParagraphFormat ParagraphFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParagraphFormat", paramsArray);
				NetOffice.WordApi.ParagraphFormat newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ParagraphFormat.LateBindingApiWrapperType) as NetOffice.WordApi.ParagraphFormat;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ParagraphFormat", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.ListFormat ListFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListFormat", paramsArray);
				NetOffice.WordApi.ListFormat newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ListFormat.LateBindingApiWrapperType) as NetOffice.WordApi.ListFormat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Bookmarks Bookmarks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Bookmarks", paramsArray);
				NetOffice.WordApi.Bookmarks newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Bookmarks.LateBindingApiWrapperType) as NetOffice.WordApi.Bookmarks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 Bold
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Bold", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Bold", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 Italic
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Italic", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Italic", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdUnderline Underline
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Underline", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdUnderline)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Underline", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdEmphasisMark EmphasisMark
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EmphasisMark", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdEmphasisMark)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EmphasisMark", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public bool DisableCharacterSpaceGrid
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisableCharacterSpaceGrid", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisableCharacterSpaceGrid", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Revisions Revisions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Revisions", paramsArray);
				NetOffice.WordApi.Revisions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Revisions.LateBindingApiWrapperType) as NetOffice.WordApi.Revisions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public object Style
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Style", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Style", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 StoryLength
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StoryLength", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdLanguageID LanguageID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LanguageID", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdLanguageID)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LanguageID", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SynonymInfo SynonymInfo
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SynonymInfo", paramsArray);
				NetOffice.WordApi.SynonymInfo newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.SynonymInfo.LateBindingApiWrapperType) as NetOffice.WordApi.SynonymInfo;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Hyperlinks Hyperlinks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hyperlinks", paramsArray);
				NetOffice.WordApi.Hyperlinks newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Hyperlinks.LateBindingApiWrapperType) as NetOffice.WordApi.Hyperlinks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.ListParagraphs ListParagraphs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListParagraphs", paramsArray);
				NetOffice.WordApi.ListParagraphs newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ListParagraphs.LateBindingApiWrapperType) as NetOffice.WordApi.ListParagraphs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Subdocuments Subdocuments
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Subdocuments", paramsArray);
				NetOffice.WordApi.Subdocuments newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Subdocuments.LateBindingApiWrapperType) as NetOffice.WordApi.Subdocuments;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public bool GrammarChecked
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GrammarChecked", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GrammarChecked", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public bool SpellingChecked
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SpellingChecked", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SpellingChecked", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdColorIndex HighlightColorIndex
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HighlightColorIndex", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdColorIndex)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HighlightColorIndex", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Columns Columns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Columns", paramsArray);
				NetOffice.WordApi.Columns newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Columns.LateBindingApiWrapperType) as NetOffice.WordApi.Columns;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Rows Rows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Rows", paramsArray);
				NetOffice.WordApi.Rows newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Rows.LateBindingApiWrapperType) as NetOffice.WordApi.Rows;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 CanEdit
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CanEdit", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 CanPaste
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CanPaste", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public bool IsEndOfRowMark
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsEndOfRowMark", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 BookmarkID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BookmarkID", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 PreviousBookmarkID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PreviousBookmarkID", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Find Find
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Find", paramsArray);
				NetOffice.WordApi.Find newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Find.LateBindingApiWrapperType) as NetOffice.WordApi.Find;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.PageSetup PageSetup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PageSetup", paramsArray);
				NetOffice.WordApi.PageSetup newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.PageSetup.LateBindingApiWrapperType) as NetOffice.WordApi.PageSetup;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PageSetup", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.ShapeRange ShapeRange
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShapeRange", paramsArray);
				NetOffice.WordApi.ShapeRange newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.WordApi.ShapeRange;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdCharacterCase Case
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Case", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdCharacterCase)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Case", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdInformation Type</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Information(NetOffice.WordApi.Enums.WdInformation type)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.PropertyGet(this, "Information", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Alias for get_Information
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdInformation Type</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public object Information(NetOffice.WordApi.Enums.WdInformation type)
		{
			return get_Information(type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.ReadabilityStatistics ReadabilityStatistics
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadabilityStatistics", paramsArray);
				NetOffice.WordApi.ReadabilityStatistics newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ReadabilityStatistics.LateBindingApiWrapperType) as NetOffice.WordApi.ReadabilityStatistics;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.ProofreadingErrors GrammaticalErrors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GrammaticalErrors", paramsArray);
				NetOffice.WordApi.ProofreadingErrors newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ProofreadingErrors.LateBindingApiWrapperType) as NetOffice.WordApi.ProofreadingErrors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.ProofreadingErrors SpellingErrors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SpellingErrors", paramsArray);
				NetOffice.WordApi.ProofreadingErrors newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ProofreadingErrors.LateBindingApiWrapperType) as NetOffice.WordApi.ProofreadingErrors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdTextOrientation Orientation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Orientation", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdTextOrientation)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Orientation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.InlineShapes InlineShapes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "InlineShapes", paramsArray);
				NetOffice.WordApi.InlineShapes newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.InlineShapes.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShapes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range NextStoryRange
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NextStoryRange", paramsArray);
				NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdLanguageID LanguageIDFarEast
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LanguageIDFarEast", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdLanguageID)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LanguageIDFarEast", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdLanguageID LanguageIDOther
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LanguageIDOther", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdLanguageID)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LanguageIDOther", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public bool LanguageDetected
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LanguageDetected", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LanguageDetected", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Single FitTextWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FitTextWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FitTextWidth", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdHorizontalInVerticalType HorizontalInVertical
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HorizontalInVertical", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdHorizontalInVerticalType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HorizontalInVertical", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdTwoLinesInOneType TwoLinesInOne
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TwoLinesInOne", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdTwoLinesInOneType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TwoLinesInOne", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public bool CombineCharacters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CombineCharacters", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CombineCharacters", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 NoProofing
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NoProofing", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "NoProofing", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Tables TopLevelTables
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TopLevelTables", paramsArray);
				NetOffice.WordApi.Tables newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Tables.LateBindingApiWrapperType) as NetOffice.WordApi.Tables;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.OfficeApi.Scripts Scripts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Scripts", paramsArray);
				NetOffice.OfficeApi.Scripts newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Scripts.LateBindingApiWrapperType) as NetOffice.OfficeApi.Scripts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdCharacterWidth CharacterWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CharacterWidth", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdCharacterWidth)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CharacterWidth", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Enums.WdKana Kana
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Kana", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdKana)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Kana", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 BoldBi
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BoldBi", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BoldBi", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 ItalicBi
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ItalicBi", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ItalicBi", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public string ID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ID", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15)]
		public NetOffice.WordApi.HTMLDivisions HTMLDivisions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HTMLDivisions", paramsArray);
				NetOffice.WordApi.HTMLDivisions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.HTMLDivisions.LateBindingApiWrapperType) as NetOffice.WordApi.HTMLDivisions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15)]
		public NetOffice.WordApi.SmartTags SmartTags
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartTags", paramsArray);
				NetOffice.WordApi.SmartTags newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.SmartTags.LateBindingApiWrapperType) as NetOffice.WordApi.SmartTags;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15)]
		public bool ShowAll
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowAll", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowAll", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15)]
		public NetOffice.WordApi.Document Document
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Document", paramsArray);
				NetOffice.WordApi.Document newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15)]
		public NetOffice.WordApi.FootnoteOptions FootnoteOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FootnoteOptions", paramsArray);
				NetOffice.WordApi.FootnoteOptions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.FootnoteOptions.LateBindingApiWrapperType) as NetOffice.WordApi.FootnoteOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15)]
		public NetOffice.WordApi.EndnoteOptions EndnoteOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EndnoteOptions", paramsArray);
				NetOffice.WordApi.EndnoteOptions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.EndnoteOptions.LateBindingApiWrapperType) as NetOffice.WordApi.EndnoteOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public NetOffice.WordApi.XMLNodes XMLNodes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLNodes", paramsArray);
				NetOffice.WordApi.XMLNodes newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType) as NetOffice.WordApi.XMLNodes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public NetOffice.WordApi.XMLNode XMLParentNode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLParentNode", paramsArray);
				NetOffice.WordApi.XMLNode newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.XMLNode.LateBindingApiWrapperType) as NetOffice.WordApi.XMLNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public NetOffice.WordApi.Editors Editors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Editors", paramsArray);
				NetOffice.WordApi.Editors newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Editors.LateBindingApiWrapperType) as NetOffice.WordApi.Editors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public string XML
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XML", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public object EnhMetaFileBits
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnhMetaFileBits", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.OMaths OMaths
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMaths", paramsArray);
				NetOffice.WordApi.OMaths newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMaths.LateBindingApiWrapperType) as NetOffice.WordApi.OMaths;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public object CharacterStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CharacterStyle", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public object ParagraphStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParagraphStyle", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public object ListStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListStyle", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public object TableStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TableStyle", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.ContentControls ContentControls
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContentControls", paramsArray);
				NetOffice.WordApi.ContentControls newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ContentControls.LateBindingApiWrapperType) as NetOffice.WordApi.ContentControls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public string WordOpenXML
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WordOpenXML", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.ContentControl ParentContentControl
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParentContentControl", paramsArray);
				NetOffice.WordApi.ContentControl newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ContentControl.LateBindingApiWrapperType) as NetOffice.WordApi.ContentControl;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15)]
		public NetOffice.WordApi.CoAuthLocks Locks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Locks", paramsArray);
				NetOffice.WordApi.CoAuthLocks newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.CoAuthLocks.LateBindingApiWrapperType) as NetOffice.WordApi.CoAuthLocks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15)]
		public NetOffice.WordApi.CoAuthUpdates Updates
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Updates", paramsArray);
				NetOffice.WordApi.CoAuthUpdates newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.CoAuthUpdates.LateBindingApiWrapperType) as NetOffice.WordApi.CoAuthUpdates;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15)]
		public NetOffice.WordApi.Conflicts Conflicts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Conflicts", paramsArray);
				NetOffice.WordApi.Conflicts newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Conflicts.LateBindingApiWrapperType) as NetOffice.WordApi.Conflicts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 15)]
		public Int32 TextVisibleOnScreen
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextVisibleOnScreen", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Select()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="start">Int32 Start</param>
		/// <param name="end">Int32 End</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SetRange(Int32 start, Int32 end)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, end);
			Invoker.Method(this, "SetRange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="direction">optional object Direction</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Collapse(object direction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(direction);
			Invoker.Method(this, "Collapse", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Collapse()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Collapse", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="text">string Text</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertBefore(string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			Invoker.Method(this, "InsertBefore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="text">string Text</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertAfter(string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			Invoker.Method(this, "InsertAfter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range Next(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "Next", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range Next()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Next", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range Next(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "Next", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range Previous(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "Previous", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range Previous()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Previous", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range Previous(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "Previous", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="extend">optional object Extend</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 StartOf(object unit, object extend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, extend);
			object returnItem = Invoker.MethodReturn(this, "StartOf", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 StartOf()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "StartOf", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 StartOf(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "StartOf", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="extend">optional object Extend</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 EndOf(object unit, object extend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, extend);
			object returnItem = Invoker.MethodReturn(this, "EndOf", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 EndOf()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "EndOf", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 EndOf(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "EndOf", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 Move(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "Move", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 Move()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Move", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 Move(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "Move", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveStart(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "MoveStart", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveStart()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "MoveStart", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveStart(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "MoveStart", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveEnd(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "MoveEnd", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveEnd()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "MoveEnd", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveEnd(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "MoveEnd", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="cset">object Cset</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveWhile(object cset, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset, count);
			object returnItem = Invoker.MethodReturn(this, "MoveWhile", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="cset">object Cset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveWhile(object cset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset);
			object returnItem = Invoker.MethodReturn(this, "MoveWhile", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="cset">object Cset</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveStartWhile(object cset, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset, count);
			object returnItem = Invoker.MethodReturn(this, "MoveStartWhile", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="cset">object Cset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveStartWhile(object cset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset);
			object returnItem = Invoker.MethodReturn(this, "MoveStartWhile", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="cset">object Cset</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveEndWhile(object cset, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset, count);
			object returnItem = Invoker.MethodReturn(this, "MoveEndWhile", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="cset">object Cset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveEndWhile(object cset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset);
			object returnItem = Invoker.MethodReturn(this, "MoveEndWhile", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="cset">object Cset</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveUntil(object cset, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset, count);
			object returnItem = Invoker.MethodReturn(this, "MoveUntil", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="cset">object Cset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveUntil(object cset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset);
			object returnItem = Invoker.MethodReturn(this, "MoveUntil", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="cset">object Cset</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveStartUntil(object cset, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset, count);
			object returnItem = Invoker.MethodReturn(this, "MoveStartUntil", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="cset">object Cset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveStartUntil(object cset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset);
			object returnItem = Invoker.MethodReturn(this, "MoveStartUntil", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="cset">object Cset</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveEndUntil(object cset, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset, count);
			object returnItem = Invoker.MethodReturn(this, "MoveEndUntil", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="cset">object Cset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 MoveEndUntil(object cset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset);
			object returnItem = Invoker.MethodReturn(this, "MoveEndUntil", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Paste()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="type">optional object Type</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertBreak(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "InsertBreak", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertBreak()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertBreak", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="range">optional object Range</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="link">optional object Link</param>
		/// <param name="attachment">optional object Attachment</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertFile(string fileName, object range, object confirmConversions, object link, object attachment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, range, confirmConversions, link, attachment);
			Invoker.Method(this, "InsertFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertFile(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "InsertFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="range">optional object Range</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertFile(string fileName, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, range);
			Invoker.Method(this, "InsertFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="range">optional object Range</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertFile(string fileName, object range, object confirmConversions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, range, confirmConversions);
			Invoker.Method(this, "InsertFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="range">optional object Range</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="link">optional object Link</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertFile(string fileName, object range, object confirmConversions, object link)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, range, confirmConversions, link);
			Invoker.Method(this, "InsertFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public bool InStory(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "InStory", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public bool InRange(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "InRange", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 Delete(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "Delete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 Delete()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Delete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 Delete(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "Delete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void WholeStory()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "WholeStory", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 Expand(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "Expand", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 Expand()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Expand", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertParagraph()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertParagraph", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertParagraphAfter()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertParagraphAfter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		/// <param name="applyFirstColumn">optional object ApplyFirstColumn</param>
		/// <param name="applyLastColumn">optional object ApplyLastColumn</param>
		/// <param name="autoFit">optional object AutoFit</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		/// <param name="applyFirstColumn">optional object ApplyFirstColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		/// <param name="applyFirstColumn">optional object ApplyFirstColumn</param>
		/// <param name="applyLastColumn">optional object ApplyLastColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		/// <param name="insertAsField">optional object InsertAsField</param>
		/// <param name="insertAsFullWidth">optional object InsertAsFullWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDateTimeOld(object dateTimeFormat, object insertAsField, object insertAsFullWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField, insertAsFullWidth);
			Invoker.Method(this, "InsertDateTimeOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDateTimeOld()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertDateTimeOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDateTimeOld(object dateTimeFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat);
			Invoker.Method(this, "InsertDateTimeOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		/// <param name="insertAsField">optional object InsertAsField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDateTimeOld(object dateTimeFormat, object insertAsField)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField);
			Invoker.Method(this, "InsertDateTimeOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="characterNumber">Int32 CharacterNumber</param>
		/// <param name="font">optional object Font</param>
		/// <param name="unicode">optional object Unicode</param>
		/// <param name="bias">optional object Bias</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertSymbol(Int32 characterNumber, object font, object unicode, object bias)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(characterNumber, font, unicode, bias);
			Invoker.Method(this, "InsertSymbol", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="characterNumber">Int32 CharacterNumber</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertSymbol(Int32 characterNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(characterNumber);
			Invoker.Method(this, "InsertSymbol", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="characterNumber">Int32 CharacterNumber</param>
		/// <param name="font">optional object Font</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertSymbol(Int32 characterNumber, object font)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(characterNumber, font);
			Invoker.Method(this, "InsertSymbol", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="characterNumber">Int32 CharacterNumber</param>
		/// <param name="font">optional object Font</param>
		/// <param name="unicode">optional object Unicode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertSymbol(Int32 characterNumber, object font, object unicode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(characterNumber, font, unicode);
			Invoker.Method(this, "InsertSymbol", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		/// <param name="insertAsHyperlink">optional object InsertAsHyperlink</param>
		/// <param name="includePosition">optional object IncludePosition</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition);
			Invoker.Method(this, "InsertCrossReference", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		/// <param name="insertAsHyperlink">optional object InsertAsHyperlink</param>
		/// <param name="includePosition">optional object IncludePosition</param>
		/// <param name="separateNumbers">optional object SeparateNumbers</param>
		/// <param name="separatorString">optional object SeparatorString</param>
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition, object separateNumbers, object separatorString)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition, separateNumbers, separatorString);
			Invoker.Method(this, "InsertCrossReference", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem);
			Invoker.Method(this, "InsertCrossReference", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		/// <param name="insertAsHyperlink">optional object InsertAsHyperlink</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem, insertAsHyperlink);
			Invoker.Method(this, "InsertCrossReference", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		/// <param name="insertAsHyperlink">optional object InsertAsHyperlink</param>
		/// <param name="includePosition">optional object IncludePosition</param>
		/// <param name="separateNumbers">optional object SeparateNumbers</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition, object separateNumbers)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition, separateNumbers);
			Invoker.Method(this, "InsertCrossReference", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		/// <param name="titleAutoText">optional object TitleAutoText</param>
		/// <param name="position">optional object Position</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertCaption(object label, object title, object titleAutoText, object position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title, titleAutoText, position);
			Invoker.Method(this, "InsertCaption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		/// <param name="titleAutoText">optional object TitleAutoText</param>
		/// <param name="position">optional object Position</param>
		/// <param name="excludeLabel">optional object ExcludeLabel</param>
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public void InsertCaption(object label, object title, object titleAutoText, object position, object excludeLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title, titleAutoText, position, excludeLabel);
			Invoker.Method(this, "InsertCaption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="label">object Label</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertCaption(object label)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label);
			Invoker.Method(this, "InsertCaption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertCaption(object label, object title)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title);
			Invoker.Method(this, "InsertCaption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		/// <param name="titleAutoText">optional object TitleAutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertCaption(object label, object title, object titleAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title, titleAutoText);
			Invoker.Method(this, "InsertCaption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CopyAsPicture()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CopyAsPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="languageID">optional object LanguageID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object languageID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, languageID);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		/// <param name="separator">optional object Separator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortAscending()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SortAscending", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SortDescending()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SortDescending", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public bool IsEqual(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "IsEqual", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Single Calculate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Calculate", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="what">optional object What</param>
		/// <param name="which">optional object Which</param>
		/// <param name="count">optional object Count</param>
		/// <param name="name">optional object Name</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range GoTo(object what, object which, object count, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, which, count, name);
			object returnItem = Invoker.MethodReturn(this, "GoTo", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range GoTo()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GoTo", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="what">optional object What</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range GoTo(object what)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what);
			object returnItem = Invoker.MethodReturn(this, "GoTo", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="what">optional object What</param>
		/// <param name="which">optional object Which</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range GoTo(object what, object which)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, which);
			object returnItem = Invoker.MethodReturn(this, "GoTo", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="what">optional object What</param>
		/// <param name="which">optional object Which</param>
		/// <param name="count">optional object Count</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range GoTo(object what, object which, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, which, count);
			object returnItem = Invoker.MethodReturn(this, "GoTo", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="what">NetOffice.WordApi.Enums.WdGoToItem What</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range GoToNext(NetOffice.WordApi.Enums.WdGoToItem what)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what);
			object returnItem = Invoker.MethodReturn(this, "GoToNext", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="what">NetOffice.WordApi.Enums.WdGoToItem What</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Range GoToPrevious(NetOffice.WordApi.Enums.WdGoToItem what)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what);
			object returnItem = Invoker.MethodReturn(this, "GoToPrevious", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="link">optional object Link</param>
		/// <param name="placement">optional object Placement</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="iconFileName">optional object IconFileName</param>
		/// <param name="iconLabel">optional object IconLabel</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType, object iconFileName, object iconLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex, link, placement, displayAsIcon, dataType, iconFileName, iconLabel);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PasteSpecial()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PasteSpecial(object iconIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="link">optional object Link</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PasteSpecial(object iconIndex, object link)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex, link);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="link">optional object Link</param>
		/// <param name="placement">optional object Placement</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PasteSpecial(object iconIndex, object link, object placement)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex, link, placement);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="link">optional object Link</param>
		/// <param name="placement">optional object Placement</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex, link, placement, displayAsIcon);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="link">optional object Link</param>
		/// <param name="placement">optional object Placement</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="dataType">optional object DataType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex, link, placement, displayAsIcon, dataType);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="link">optional object Link</param>
		/// <param name="placement">optional object Placement</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="iconFileName">optional object IconFileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType, object iconFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex, link, placement, displayAsIcon, dataType, iconFileName);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void LookupNameProperties()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "LookupNameProperties", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="statistic">NetOffice.WordApi.Enums.WdStatistic Statistic</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public Int32 ComputeStatistics(NetOffice.WordApi.Enums.WdStatistic statistic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(statistic);
			object returnItem = Invoker.MethodReturn(this, "ComputeStatistics", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="direction">Int32 Direction</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Relocate(Int32 direction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(direction);
			Invoker.Method(this, "Relocate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSynonyms()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckSynonyms", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="edition">string Edition</param>
		/// <param name="format">optional object Format</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SubscribeTo(string edition, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, format);
			Invoker.Method(this, "SubscribeTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="edition">string Edition</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void SubscribeTo(string edition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition);
			Invoker.Method(this, "SubscribeTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="containsPICT">optional object ContainsPICT</param>
		/// <param name="containsRTF">optional object ContainsRTF</param>
		/// <param name="containsText">optional object ContainsText</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CreatePublisher(object edition, object containsPICT, object containsRTF, object containsText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, containsPICT, containsRTF, containsText);
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CreatePublisher()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CreatePublisher(object edition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition);
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="containsPICT">optional object ContainsPICT</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CreatePublisher(object edition, object containsPICT)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, containsPICT);
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="containsPICT">optional object ContainsPICT</param>
		/// <param name="containsRTF">optional object ContainsRTF</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CreatePublisher(object edition, object containsPICT, object containsRTF)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, containsPICT, containsRTF);
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertAutoText()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertAutoText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="dataSource">optional object DataSource</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="includeFields">optional object IncludeFields</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource, object from, object to, object includeFields)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate, dataSource, from, to, includeFields);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style, object linkToSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style, linkToSource);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="connection">optional object Connection</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style, linkToSource, connection);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style, linkToSource, connection, sQLStatement);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style, linkToSource, connection, sQLStatement, sQLStatement1);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="dataSource">optional object DataSource</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate, dataSource);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="dataSource">optional object DataSource</param>
		/// <param name="from">optional object From</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource, object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate, dataSource, from);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="style">optional object Style</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="dataSource">optional object DataSource</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource, object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate, dataSource, from, to);
			Invoker.Method(this, "InsertDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void AutoFormat()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckGrammar()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckGrammar", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		/// <param name="customDictionary8">optional object CustomDictionary8</param>
		/// <param name="customDictionary9">optional object CustomDictionary9</param>
		/// <param name="customDictionary10">optional object CustomDictionary10</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling(object customDictionary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		/// <param name="customDictionary8">optional object CustomDictionary8</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		/// <param name="customDictionary8">optional object CustomDictionary8</param>
		/// <param name="customDictionary9">optional object CustomDictionary9</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		/// <param name="customDictionary8">optional object CustomDictionary8</param>
		/// <param name="customDictionary9">optional object CustomDictionary9</param>
		/// <param name="customDictionary10">optional object CustomDictionary10</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, mainDictionary);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, mainDictionary, suggestionMode);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		/// <param name="customDictionary8">optional object CustomDictionary8</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		/// <param name="customDictionary8">optional object CustomDictionary8</param>
		/// <param name="customDictionary9">optional object CustomDictionary9</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertParagraphBefore()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertParagraphBefore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void NextSubdocument()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "NextSubdocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PreviousSubdocument()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PreviousSubdocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="conversionsMode">optional object ConversionsMode</param>
		/// <param name="fastConversion">optional object FastConversion</param>
		/// <param name="checkHangulEnding">optional object CheckHangulEnding</param>
		/// <param name="enableRecentOrdering">optional object EnableRecentOrdering</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void ConvertHangulAndHanja(object conversionsMode, object fastConversion, object checkHangulEnding, object enableRecentOrdering, object customDictionary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(conversionsMode, fastConversion, checkHangulEnding, enableRecentOrdering, customDictionary);
			Invoker.Method(this, "ConvertHangulAndHanja", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void ConvertHangulAndHanja()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ConvertHangulAndHanja", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="conversionsMode">optional object ConversionsMode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void ConvertHangulAndHanja(object conversionsMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(conversionsMode);
			Invoker.Method(this, "ConvertHangulAndHanja", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="conversionsMode">optional object ConversionsMode</param>
		/// <param name="fastConversion">optional object FastConversion</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void ConvertHangulAndHanja(object conversionsMode, object fastConversion)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(conversionsMode, fastConversion);
			Invoker.Method(this, "ConvertHangulAndHanja", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="conversionsMode">optional object ConversionsMode</param>
		/// <param name="fastConversion">optional object FastConversion</param>
		/// <param name="checkHangulEnding">optional object CheckHangulEnding</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void ConvertHangulAndHanja(object conversionsMode, object fastConversion, object checkHangulEnding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(conversionsMode, fastConversion, checkHangulEnding);
			Invoker.Method(this, "ConvertHangulAndHanja", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="conversionsMode">optional object ConversionsMode</param>
		/// <param name="fastConversion">optional object FastConversion</param>
		/// <param name="checkHangulEnding">optional object CheckHangulEnding</param>
		/// <param name="enableRecentOrdering">optional object EnableRecentOrdering</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void ConvertHangulAndHanja(object conversionsMode, object fastConversion, object checkHangulEnding, object enableRecentOrdering)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(conversionsMode, fastConversion, checkHangulEnding, enableRecentOrdering);
			Invoker.Method(this, "ConvertHangulAndHanja", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PasteAsNestedTable()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PasteAsNestedTable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="style">object Style</param>
		/// <param name="symbol">optional object Symbol</param>
		/// <param name="enclosedText">optional object EnclosedText</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void ModifyEnclosure(object style, object symbol, object enclosedText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, symbol, enclosedText);
			Invoker.Method(this, "ModifyEnclosure", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="style">object Style</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void ModifyEnclosure(object style)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style);
			Invoker.Method(this, "ModifyEnclosure", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="style">object Style</param>
		/// <param name="symbol">optional object Symbol</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void ModifyEnclosure(object style, object symbol)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, symbol);
			Invoker.Method(this, "ModifyEnclosure", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
		/// <param name="raise">optional Int32 Raise = 0</param>
		/// <param name="fontSize">optional Int32 FontSize = 0</param>
		/// <param name="fontName">optional string FontName = </param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PhoneticGuide(string text, NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType alignment, Int32 raise, Int32 fontSize, string fontName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, alignment, raise, fontSize, fontName);
			Invoker.Method(this, "PhoneticGuide", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="text">string Text</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PhoneticGuide(string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			Invoker.Method(this, "PhoneticGuide", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PhoneticGuide(string text, NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType alignment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, alignment);
			Invoker.Method(this, "PhoneticGuide", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
		/// <param name="raise">optional Int32 Raise = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PhoneticGuide(string text, NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType alignment, Int32 raise)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, alignment, raise);
			Invoker.Method(this, "PhoneticGuide", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
		/// <param name="raise">optional Int32 Raise = 0</param>
		/// <param name="fontSize">optional Int32 FontSize = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void PhoneticGuide(string text, NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType alignment, Int32 raise, Int32 fontSize)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, alignment, raise, fontSize);
			Invoker.Method(this, "PhoneticGuide", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		/// <param name="insertAsField">optional object InsertAsField</param>
		/// <param name="insertAsFullWidth">optional object InsertAsFullWidth</param>
		/// <param name="dateLanguage">optional object DateLanguage</param>
		/// <param name="calendarType">optional object CalendarType</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth, object dateLanguage, object calendarType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField, insertAsFullWidth, dateLanguage, calendarType);
			Invoker.Method(this, "InsertDateTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDateTime()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertDateTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDateTime(object dateTimeFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat);
			Invoker.Method(this, "InsertDateTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		/// <param name="insertAsField">optional object InsertAsField</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDateTime(object dateTimeFormat, object insertAsField)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField);
			Invoker.Method(this, "InsertDateTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		/// <param name="insertAsField">optional object InsertAsField</param>
		/// <param name="insertAsFullWidth">optional object InsertAsFullWidth</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField, insertAsFullWidth);
			Invoker.Method(this, "InsertDateTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		/// <param name="insertAsField">optional object InsertAsField</param>
		/// <param name="insertAsFullWidth">optional object InsertAsFullWidth</param>
		/// <param name="dateLanguage">optional object DateLanguage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth, object dateLanguage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField, insertAsFullWidth, dateLanguage);
			Invoker.Method(this, "InsertDateTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		/// <param name="ignoreDiacritics">optional object IgnoreDiacritics</param>
		/// <param name="ignoreHe">optional object IgnoreHe</param>
		/// <param name="languageID">optional object LanguageID</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		/// <param name="separator">optional object Separator</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		/// <param name="ignoreDiacritics">optional object IgnoreDiacritics</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="sortColumn">optional object SortColumn</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		/// <param name="ignoreDiacritics">optional object IgnoreDiacritics</param>
		/// <param name="ignoreHe">optional object IgnoreHe</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void DetectLanguage()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DetectLanguage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		/// <param name="applyFirstColumn">optional object ApplyFirstColumn</param>
		/// <param name="applyLastColumn">optional object ApplyLastColumn</param>
		/// <param name="autoFit">optional object AutoFit</param>
		/// <param name="autoFitBehavior">optional object AutoFitBehavior</param>
		/// <param name="defaultTableBehavior">optional object DefaultTableBehavior</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit, object autoFitBehavior, object defaultTableBehavior)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit, autoFitBehavior, defaultTableBehavior);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		/// <param name="applyFirstColumn">optional object ApplyFirstColumn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		/// <param name="applyFirstColumn">optional object ApplyFirstColumn</param>
		/// <param name="applyLastColumn">optional object ApplyLastColumn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		/// <param name="applyFirstColumn">optional object ApplyFirstColumn</param>
		/// <param name="applyLastColumn">optional object ApplyLastColumn</param>
		/// <param name="autoFit">optional object AutoFit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		/// <param name="applyFirstColumn">optional object ApplyFirstColumn</param>
		/// <param name="applyLastColumn">optional object ApplyLastColumn</param>
		/// <param name="autoFit">optional object AutoFit</param>
		/// <param name="autoFitBehavior">optional object AutoFitBehavior</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit, object autoFitBehavior)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit, autoFitBehavior);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="wdTCSCConverterDirection">optional NetOffice.WordApi.Enums.WdTCSCConverterDirection WdTCSCConverterDirection = 2</param>
		/// <param name="commonTerms">optional bool CommonTerms = false</param>
		/// <param name="useVariants">optional bool UseVariants = false</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void TCSCConverter(NetOffice.WordApi.Enums.WdTCSCConverterDirection wdTCSCConverterDirection, bool commonTerms, bool useVariants)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wdTCSCConverterDirection, commonTerms, useVariants);
			Invoker.Method(this, "TCSCConverter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void TCSCConverter()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "TCSCConverter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="wdTCSCConverterDirection">optional NetOffice.WordApi.Enums.WdTCSCConverterDirection WdTCSCConverterDirection = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void TCSCConverter(NetOffice.WordApi.Enums.WdTCSCConverterDirection wdTCSCConverterDirection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wdTCSCConverterDirection);
			Invoker.Method(this, "TCSCConverter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="wdTCSCConverterDirection">optional NetOffice.WordApi.Enums.WdTCSCConverterDirection WdTCSCConverterDirection = 2</param>
		/// <param name="commonTerms">optional bool CommonTerms = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		public void TCSCConverter(NetOffice.WordApi.Enums.WdTCSCConverterDirection wdTCSCConverterDirection, bool commonTerms)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wdTCSCConverterDirection, commonTerms);
			Invoker.Method(this, "TCSCConverter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdRecoveryType Type</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15)]
		public void PasteAndFormat(NetOffice.WordApi.Enums.WdRecoveryType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "PasteAndFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="linkedToExcel">bool LinkedToExcel</param>
		/// <param name="wordFormatting">bool WordFormatting</param>
		/// <param name="rTF">bool RTF</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15)]
		public void PasteExcelTable(bool linkedToExcel, bool wordFormatting, bool rTF)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(linkedToExcel, wordFormatting, rTF);
			Invoker.Method(this, "PasteExcelTable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15)]
		public void PasteAppendTable()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PasteAppendTable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		/// <param name="insertAsHyperlink">optional object InsertAsHyperlink</param>
		/// <param name="includePosition">optional object IncludePosition</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition);
			Invoker.Method(this, "InsertCrossReference_2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem);
			Invoker.Method(this, "InsertCrossReference_2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		/// <param name="insertAsHyperlink">optional object InsertAsHyperlink</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem, insertAsHyperlink);
			Invoker.Method(this, "InsertCrossReference_2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		/// <param name="titleAutoText">optional object TitleAutoText</param>
		/// <param name="position">optional object Position</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public void InsertCaptionXP(object label, object title, object titleAutoText, object position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title, titleAutoText, position);
			Invoker.Method(this, "InsertCaptionXP", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="label">object Label</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public void InsertCaptionXP(object label)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label);
			Invoker.Method(this, "InsertCaptionXP", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public void InsertCaptionXP(object label, object title)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title);
			Invoker.Method(this, "InsertCaptionXP", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		/// <param name="titleAutoText">optional object TitleAutoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public void InsertCaptionXP(object label, object title, object titleAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title, titleAutoText);
			Invoker.Method(this, "InsertCaptionXP", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="editorID">optional object EditorID</param>
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public NetOffice.WordApi.Range GoToEditableRange(object editorID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(editorID);
			object returnItem = Invoker.MethodReturn(this, "GoToEditableRange", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public NetOffice.WordApi.Range GoToEditableRange()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GoToEditableRange", paramsArray);
			NetOffice.WordApi.Range newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="xML">string XML</param>
		/// <param name="transform">optional object Transform</param>
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public void InsertXML(string xML, object transform)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xML, transform);
			Invoker.Method(this, "InsertXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15
		/// </summary>
		/// <param name="xML">string XML</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15)]
		public void InsertXML(string xML)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xML);
			Invoker.Method(this, "InsertXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="format">NetOffice.WordApi.Enums.WdSaveFormat Format</param>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportFragment(string fileName, NetOffice.WordApi.Enums.WdSaveFormat format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, format);
			Invoker.Method(this, "ExportFragment", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="level">Int16 Level</param>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void SetListLevel(Int16 level)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(level);
			Invoker.Method(this, "SetListLevel", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="alignment">Int32 Alignment</param>
		/// <param name="relativeTo">optional Int32 RelativeTo = 0</param>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void InsertAlignmentTab(Int32 alignment, Int32 relativeTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(alignment, relativeTo);
			Invoker.Method(this, "InsertAlignmentTab", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="alignment">Int32 Alignment</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void InsertAlignmentTab(Int32 alignment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(alignment);
			Invoker.Method(this, "InsertAlignmentTab", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="matchDestination">optional bool MatchDestination = false</param>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ImportFragment(string fileName, bool matchDestination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, matchDestination);
			Invoker.Method(this, "ImportFragment", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ImportFragment(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "ImportFragment", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="fixedFormatExtClassPtr">optional object FixedFormatExtClassPtr</param>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, bool openAfterExport, NetOffice.WordApi.Enums.WdExportOptimizeFor optimizeFor, bool exportCurrentPage, NetOffice.WordApi.Enums.WdExportItem item, bool includeDocProps, bool keepIRM, NetOffice.WordApi.Enums.WdExportCreateBookmarks createBookmarks, bool docStructureTags, bool bitmapMissingFonts, bool useISO19005_1, object fixedFormatExtClassPtr)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1, fixedFormatExtClassPtr);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, bool openAfterExport)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, bool openAfterExport, NetOffice.WordApi.Enums.WdExportOptimizeFor optimizeFor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, bool openAfterExport, NetOffice.WordApi.Enums.WdExportOptimizeFor optimizeFor, bool exportCurrentPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, bool openAfterExport, NetOffice.WordApi.Enums.WdExportOptimizeFor optimizeFor, bool exportCurrentPage, NetOffice.WordApi.Enums.WdExportItem item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, bool openAfterExport, NetOffice.WordApi.Enums.WdExportOptimizeFor optimizeFor, bool exportCurrentPage, NetOffice.WordApi.Enums.WdExportItem item, bool includeDocProps)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, bool openAfterExport, NetOffice.WordApi.Enums.WdExportOptimizeFor optimizeFor, bool exportCurrentPage, NetOffice.WordApi.Enums.WdExportItem item, bool includeDocProps, bool keepIRM)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, bool openAfterExport, NetOffice.WordApi.Enums.WdExportOptimizeFor optimizeFor, bool exportCurrentPage, NetOffice.WordApi.Enums.WdExportItem item, bool includeDocProps, bool keepIRM, NetOffice.WordApi.Enums.WdExportCreateBookmarks createBookmarks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, bool openAfterExport, NetOffice.WordApi.Enums.WdExportOptimizeFor optimizeFor, bool exportCurrentPage, NetOffice.WordApi.Enums.WdExportItem item, bool includeDocProps, bool keepIRM, NetOffice.WordApi.Enums.WdExportCreateBookmarks createBookmarks, bool docStructureTags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, bool openAfterExport, NetOffice.WordApi.Enums.WdExportOptimizeFor optimizeFor, bool exportCurrentPage, NetOffice.WordApi.Enums.WdExportItem item, bool includeDocProps, bool keepIRM, NetOffice.WordApi.Enums.WdExportCreateBookmarks createBookmarks, bool docStructureTags, bool bitmapMissingFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, bool openAfterExport, NetOffice.WordApi.Enums.WdExportOptimizeFor optimizeFor, bool exportCurrentPage, NetOffice.WordApi.Enums.WdExportItem item, bool includeDocProps, bool keepIRM, NetOffice.WordApi.Enums.WdExportCreateBookmarks createBookmarks, bool docStructureTags, bool bitmapMissingFonts, bool useISO19005_1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		/// <param name="ignoreDiacritics">optional object IgnoreDiacritics</param>
		/// <param name="ignoreHe">optional object IgnoreHe</param>
		/// <param name="languageID">optional object LanguageID</param>
		[SupportByVersionAttribute("Word", 15)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15)]
		public void SortByHeadings()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15)]
		public void SortByHeadings(object sortFieldType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15)]
		public void SortByHeadings(object sortFieldType, object sortOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive, bidiSort);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		/// <param name="ignoreDiacritics">optional object IgnoreDiacritics</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		/// <param name="ignoreDiacritics">optional object IgnoreDiacritics</param>
		/// <param name="ignoreHe">optional object IgnoreHe</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}