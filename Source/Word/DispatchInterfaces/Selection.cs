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
	/// Selection
	///</summary>
	public class Selection_ : COMObject
	{
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Selection_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection_() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="dataOnly">optional bool DataOnly</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838928.aspx
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_XML(object dataOnly)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(dataOnly);
			object returnItem = Invoker.PropertyGet(this, "XML", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838928.aspx
		/// Alias for get_XML
		/// </summary>
		/// <param name="dataOnly">optional bool DataOnly</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public string XML(object dataOnly)
		{
			return get_XML(dataOnly);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// DispatchInterface Selection 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821411.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Selection : Selection_
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
                    _type = typeof(Selection);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Selection(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192754.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836670.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range FormattedText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormattedText", paramsArray);
				NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormattedText", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839485.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834869.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837859.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Font Font
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Font", paramsArray);
				NetOffice.WordApi.Font newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Font.LateBindingApiWrapperType) as NetOffice.WordApi.Font;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Font", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821048.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSelectionType Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdSelectionType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191739.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838978.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Style
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Style", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845908.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Tables Tables
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Tables", paramsArray);
				NetOffice.WordApi.Tables newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Tables.LateBindingApiWrapperType) as NetOffice.WordApi.Tables;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837460.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Words Words
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Words", paramsArray);
				NetOffice.WordApi.Words newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Words.LateBindingApiWrapperType) as NetOffice.WordApi.Words;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193720.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Sentences Sentences
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Sentences", paramsArray);
				NetOffice.WordApi.Sentences newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Sentences.LateBindingApiWrapperType) as NetOffice.WordApi.Sentences;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196946.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Characters Characters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Characters", paramsArray);
				NetOffice.WordApi.Characters newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Characters.LateBindingApiWrapperType) as NetOffice.WordApi.Characters;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197009.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Footnotes Footnotes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Footnotes", paramsArray);
				NetOffice.WordApi.Footnotes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Footnotes.LateBindingApiWrapperType) as NetOffice.WordApi.Footnotes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841006.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Endnotes Endnotes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Endnotes", paramsArray);
				NetOffice.WordApi.Endnotes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Endnotes.LateBindingApiWrapperType) as NetOffice.WordApi.Endnotes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823219.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Comments Comments
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Comments", paramsArray);
				NetOffice.WordApi.Comments newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Comments.LateBindingApiWrapperType) as NetOffice.WordApi.Comments;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195296.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Cells Cells
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
				NetOffice.WordApi.Cells newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Cells.LateBindingApiWrapperType) as NetOffice.WordApi.Cells;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836277.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Sections Sections
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Sections", paramsArray);
				NetOffice.WordApi.Sections newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Sections.LateBindingApiWrapperType) as NetOffice.WordApi.Sections;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840393.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Paragraphs Paragraphs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Paragraphs", paramsArray);
				NetOffice.WordApi.Paragraphs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Paragraphs.LateBindingApiWrapperType) as NetOffice.WordApi.Paragraphs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193012.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Borders Borders
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Borders", paramsArray);
				NetOffice.WordApi.Borders newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Borders.LateBindingApiWrapperType) as NetOffice.WordApi.Borders;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Borders", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192021.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shading Shading
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shading", paramsArray);
				NetOffice.WordApi.Shading newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Shading.LateBindingApiWrapperType) as NetOffice.WordApi.Shading;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845839.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Fields Fields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Fields", paramsArray);
				NetOffice.WordApi.Fields newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Fields.LateBindingApiWrapperType) as NetOffice.WordApi.Fields;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838906.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FormFields FormFields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormFields", paramsArray);
				NetOffice.WordApi.FormFields newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.FormFields.LateBindingApiWrapperType) as NetOffice.WordApi.FormFields;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838307.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Frames Frames
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Frames", paramsArray);
				NetOffice.WordApi.Frames newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Frames.LateBindingApiWrapperType) as NetOffice.WordApi.Frames;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193858.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ParagraphFormat ParagraphFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParagraphFormat", paramsArray);
				NetOffice.WordApi.ParagraphFormat newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ParagraphFormat.LateBindingApiWrapperType) as NetOffice.WordApi.ParagraphFormat;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ParagraphFormat", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197430.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.PageSetup PageSetup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PageSetup", paramsArray);
				NetOffice.WordApi.PageSetup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.PageSetup.LateBindingApiWrapperType) as NetOffice.WordApi.PageSetup;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PageSetup", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193356.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Bookmarks Bookmarks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Bookmarks", paramsArray);
				NetOffice.WordApi.Bookmarks newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Bookmarks.LateBindingApiWrapperType) as NetOffice.WordApi.Bookmarks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836357.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838983.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196398.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191830.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838134.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlinks Hyperlinks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hyperlinks", paramsArray);
				NetOffice.WordApi.Hyperlinks newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Hyperlinks.LateBindingApiWrapperType) as NetOffice.WordApi.Hyperlinks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194663.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Columns Columns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Columns", paramsArray);
				NetOffice.WordApi.Columns newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Columns.LateBindingApiWrapperType) as NetOffice.WordApi.Columns;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821842.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Rows Rows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Rows", paramsArray);
				NetOffice.WordApi.Rows newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Rows.LateBindingApiWrapperType) as NetOffice.WordApi.Rows;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836744.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.HeaderFooter HeaderFooter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HeaderFooter", paramsArray);
				NetOffice.WordApi.HeaderFooter newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.HeaderFooter.LateBindingApiWrapperType) as NetOffice.WordApi.HeaderFooter;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845161.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840519.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193388.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197434.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Find Find
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Find", paramsArray);
				NetOffice.WordApi.Find newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Find.LateBindingApiWrapperType) as NetOffice.WordApi.Find;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845594.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Range
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
				NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820800.aspx
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdInformation Type</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Information(NetOffice.WordApi.Enums.WdInformation type)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.PropertyGet(this, "Information", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820800.aspx
		/// Alias for get_Information
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdInformation Type</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Information(NetOffice.WordApi.Enums.WdInformation type)
		{
			return get_Information(type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837479.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSelectionFlags Flags
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Flags", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdSelectionFlags)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Flags", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835497.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Active
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Active", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820824.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool StartIsActive
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StartIsActive", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "StartIsActive", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822970.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool IPAtEndOfLine
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IPAtEndOfLine", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821400.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool ExtendMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ExtendMode", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ExtendMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839310.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool ColumnSelectMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColumnSelectMode", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ColumnSelectMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821992.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193084.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShapes InlineShapes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "InlineShapes", paramsArray);
				NetOffice.WordApi.InlineShapes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.InlineShapes.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShapes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192167.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196980.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839166.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844964.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Document
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Document", paramsArray);
				NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836759.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ShapeRange ShapeRange
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShapeRange", paramsArray);
				NetOffice.WordApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.WordApi.ShapeRange;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196937.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821380.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Tables TopLevelTables
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TopLevelTables", paramsArray);
				NetOffice.WordApi.Tables newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Tables.LateBindingApiWrapperType) as NetOffice.WordApi.Tables;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192601.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821699.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198226.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.HTMLDivisions HTMLDivisions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HTMLDivisions", paramsArray);
				NetOffice.WordApi.HTMLDivisions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.HTMLDivisions.LateBindingApiWrapperType) as NetOffice.WordApi.HTMLDivisions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.SmartTags SmartTags
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartTags", paramsArray);
				NetOffice.WordApi.SmartTags newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.SmartTags.LateBindingApiWrapperType) as NetOffice.WordApi.SmartTags;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191940.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.ShapeRange ChildShapeRange
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChildShapeRange", paramsArray);
				NetOffice.WordApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.WordApi.ShapeRange;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191804.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool HasChildShapeRange
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasChildShapeRange", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845098.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.FootnoteOptions FootnoteOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FootnoteOptions", paramsArray);
				NetOffice.WordApi.FootnoteOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.FootnoteOptions.LateBindingApiWrapperType) as NetOffice.WordApi.FootnoteOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192368.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.EndnoteOptions EndnoteOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EndnoteOptions", paramsArray);
				NetOffice.WordApi.EndnoteOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.EndnoteOptions.LateBindingApiWrapperType) as NetOffice.WordApi.EndnoteOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes XMLNodes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLNodes", paramsArray);
				NetOffice.WordApi.XMLNodes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType) as NetOffice.WordApi.XMLNodes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode XMLParentNode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLParentNode", paramsArray);
				NetOffice.WordApi.XMLNode newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.XMLNode.LateBindingApiWrapperType) as NetOffice.WordApi.XMLNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837314.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Editors Editors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Editors", paramsArray);
				NetOffice.WordApi.Editors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Editors.LateBindingApiWrapperType) as NetOffice.WordApi.Editors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838928.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
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
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840039.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public object EnhMetaFileBits
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnhMetaFileBits", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838161.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMaths OMaths
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMaths", paramsArray);
				NetOffice.WordApi.OMaths newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMaths.LateBindingApiWrapperType) as NetOffice.WordApi.OMaths;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820971.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ContentControls ContentControls
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContentControls", paramsArray);
				NetOffice.WordApi.ContentControls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ContentControls.LateBindingApiWrapperType) as NetOffice.WordApi.ContentControls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ContentControl ParentContentControl
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParentContentControl", paramsArray);
				NetOffice.WordApi.ContentControl newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ContentControl.LateBindingApiWrapperType) as NetOffice.WordApi.ContentControl;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845714.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Select()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192352.aspx
		/// </summary>
		/// <param name="start">Int32 Start</param>
		/// <param name="end">Int32 End</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SetRange(Int32 start, Int32 end)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, end);
			Invoker.Method(this, "SetRange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834294.aspx
		/// </summary>
		/// <param name="direction">optional object Direction</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Collapse(object direction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(direction);
			Invoker.Method(this, "Collapse", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834294.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Collapse()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Collapse", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845077.aspx
		/// </summary>
		/// <param name="text">string Text</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertBefore(string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			Invoker.Method(this, "InsertBefore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192184.aspx
		/// </summary>
		/// <param name="text">string Text</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertAfter(string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			Invoker.Method(this, "InsertAfter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195124.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Next(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "Next", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195124.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Next()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Next", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195124.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Next(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "Next", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822303.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Previous(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "Previous", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822303.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Previous()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Previous", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822303.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Previous(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "Previous", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196209.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="extend">optional object Extend</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 StartOf(object unit, object extend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, extend);
			object returnItem = Invoker.MethodReturn(this, "StartOf", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196209.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 StartOf()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "StartOf", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196209.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 StartOf(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "StartOf", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193383.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="extend">optional object Extend</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 EndOf(object unit, object extend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, extend);
			object returnItem = Invoker.MethodReturn(this, "EndOf", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193383.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 EndOf()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "EndOf", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193383.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 EndOf(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "EndOf", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822886.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Move(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "Move", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822886.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Move()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Move", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822886.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Move(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "Move", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837936.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStart(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "MoveStart", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837936.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStart()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "MoveStart", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837936.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStart(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "MoveStart", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845693.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEnd(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "MoveEnd", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845693.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEnd()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "MoveEnd", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845693.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEnd(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "MoveEnd", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837303.aspx
		/// </summary>
		/// <param name="cset">object Cset</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveWhile(object cset, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset, count);
			object returnItem = Invoker.MethodReturn(this, "MoveWhile", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837303.aspx
		/// </summary>
		/// <param name="cset">object Cset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveWhile(object cset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset);
			object returnItem = Invoker.MethodReturn(this, "MoveWhile", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837161.aspx
		/// </summary>
		/// <param name="cset">object Cset</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStartWhile(object cset, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset, count);
			object returnItem = Invoker.MethodReturn(this, "MoveStartWhile", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837161.aspx
		/// </summary>
		/// <param name="cset">object Cset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStartWhile(object cset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset);
			object returnItem = Invoker.MethodReturn(this, "MoveStartWhile", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837730.aspx
		/// </summary>
		/// <param name="cset">object Cset</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEndWhile(object cset, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset, count);
			object returnItem = Invoker.MethodReturn(this, "MoveEndWhile", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837730.aspx
		/// </summary>
		/// <param name="cset">object Cset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEndWhile(object cset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset);
			object returnItem = Invoker.MethodReturn(this, "MoveEndWhile", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822578.aspx
		/// </summary>
		/// <param name="cset">object Cset</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveUntil(object cset, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset, count);
			object returnItem = Invoker.MethodReturn(this, "MoveUntil", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822578.aspx
		/// </summary>
		/// <param name="cset">object Cset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveUntil(object cset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset);
			object returnItem = Invoker.MethodReturn(this, "MoveUntil", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835726.aspx
		/// </summary>
		/// <param name="cset">object Cset</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStartUntil(object cset, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset, count);
			object returnItem = Invoker.MethodReturn(this, "MoveStartUntil", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835726.aspx
		/// </summary>
		/// <param name="cset">object Cset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStartUntil(object cset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset);
			object returnItem = Invoker.MethodReturn(this, "MoveStartUntil", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839831.aspx
		/// </summary>
		/// <param name="cset">object Cset</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEndUntil(object cset, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset, count);
			object returnItem = Invoker.MethodReturn(this, "MoveEndUntil", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839831.aspx
		/// </summary>
		/// <param name="cset">object Cset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEndUntil(object cset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cset);
			object returnItem = Invoker.MethodReturn(this, "MoveEndUntil", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192037.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196538.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840284.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Paste()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192797.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertBreak(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "InsertBreak", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192797.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertBreak()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertBreak", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="range">optional object Range</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="link">optional object Link</param>
		/// <param name="attachment">optional object Attachment</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertFile(string fileName, object range, object confirmConversions, object link, object attachment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, range, confirmConversions, link, attachment);
			Invoker.Method(this, "InsertFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertFile(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "InsertFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="range">optional object Range</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertFile(string fileName, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, range);
			Invoker.Method(this, "InsertFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="range">optional object Range</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertFile(string fileName, object range, object confirmConversions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, range, confirmConversions);
			Invoker.Method(this, "InsertFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="range">optional object Range</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="link">optional object Link</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertFile(string fileName, object range, object confirmConversions, object link)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, range, confirmConversions, link);
			Invoker.Method(this, "InsertFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192633.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool InStory(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "InStory", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193660.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool InRange(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "InRange", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193432.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Delete(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "Delete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193432.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Delete()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Delete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193432.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Delete(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "Delete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822873.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Expand(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "Expand", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822873.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Expand()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Expand", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837485.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertParagraph()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertParagraph", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836408.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertParagraphAfter()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertParagraphAfter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTableOld", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		/// <param name="insertAsField">optional object InsertAsField</param>
		/// <param name="insertAsFullWidth">optional object InsertAsFullWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTimeOld(object dateTimeFormat, object insertAsField, object insertAsFullWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField, insertAsFullWidth);
			Invoker.Method(this, "InsertDateTimeOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTimeOld()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertDateTimeOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTimeOld(object dateTimeFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat);
			Invoker.Method(this, "InsertDateTimeOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		/// <param name="insertAsField">optional object InsertAsField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTimeOld(object dateTimeFormat, object insertAsField)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField);
			Invoker.Method(this, "InsertDateTimeOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx
		/// </summary>
		/// <param name="characterNumber">Int32 CharacterNumber</param>
		/// <param name="font">optional object Font</param>
		/// <param name="unicode">optional object Unicode</param>
		/// <param name="bias">optional object Bias</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertSymbol(Int32 characterNumber, object font, object unicode, object bias)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(characterNumber, font, unicode, bias);
			Invoker.Method(this, "InsertSymbol", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx
		/// </summary>
		/// <param name="characterNumber">Int32 CharacterNumber</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertSymbol(Int32 characterNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(characterNumber);
			Invoker.Method(this, "InsertSymbol", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx
		/// </summary>
		/// <param name="characterNumber">Int32 CharacterNumber</param>
		/// <param name="font">optional object Font</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertSymbol(Int32 characterNumber, object font)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(characterNumber, font);
			Invoker.Method(this, "InsertSymbol", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx
		/// </summary>
		/// <param name="characterNumber">Int32 CharacterNumber</param>
		/// <param name="font">optional object Font</param>
		/// <param name="unicode">optional object Unicode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertSymbol(Int32 characterNumber, object font, object unicode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(characterNumber, font, unicode);
			Invoker.Method(this, "InsertSymbol", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		/// <param name="insertAsHyperlink">optional object InsertAsHyperlink</param>
		/// <param name="includePosition">optional object IncludePosition</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition);
			Invoker.Method(this, "InsertCrossReference", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		/// <param name="insertAsHyperlink">optional object InsertAsHyperlink</param>
		/// <param name="includePosition">optional object IncludePosition</param>
		/// <param name="separateNumbers">optional object SeparateNumbers</param>
		/// <param name="separatorString">optional object SeparatorString</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition, object separateNumbers, object separatorString)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition, separateNumbers, separatorString);
			Invoker.Method(this, "InsertCrossReference", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem);
			Invoker.Method(this, "InsertCrossReference", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		/// <param name="insertAsHyperlink">optional object InsertAsHyperlink</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem, insertAsHyperlink);
			Invoker.Method(this, "InsertCrossReference", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		/// <param name="insertAsHyperlink">optional object InsertAsHyperlink</param>
		/// <param name="includePosition">optional object IncludePosition</param>
		/// <param name="separateNumbers">optional object SeparateNumbers</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition, object separateNumbers)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition, separateNumbers);
			Invoker.Method(this, "InsertCrossReference", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		/// <param name="titleAutoText">optional object TitleAutoText</param>
		/// <param name="position">optional object Position</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertCaption(object label, object title, object titleAutoText, object position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title, titleAutoText, position);
			Invoker.Method(this, "InsertCaption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		/// <param name="titleAutoText">optional object TitleAutoText</param>
		/// <param name="position">optional object Position</param>
		/// <param name="excludeLabel">optional object ExcludeLabel</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void InsertCaption(object label, object title, object titleAutoText, object position, object excludeLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title, titleAutoText, position, excludeLabel);
			Invoker.Method(this, "InsertCaption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx
		/// </summary>
		/// <param name="label">object Label</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertCaption(object label)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label);
			Invoker.Method(this, "InsertCaption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertCaption(object label, object title)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title);
			Invoker.Method(this, "InsertCaption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		/// <param name="titleAutoText">optional object TitleAutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertCaption(object label, object title, object titleAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title, titleAutoText);
			Invoker.Method(this, "InsertCaption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840576.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CopyAsPicture()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CopyAsPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object languageID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, languageID);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821863.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortAscending()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SortAscending", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845052.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortDescending()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SortDescending", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196258.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool IsEqual(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "IsEqual", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835748.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single Calculate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Calculate", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx
		/// </summary>
		/// <param name="what">optional object What</param>
		/// <param name="which">optional object Which</param>
		/// <param name="count">optional object Count</param>
		/// <param name="name">optional object Name</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo(object what, object which, object count, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, which, count, name);
			object returnItem = Invoker.MethodReturn(this, "GoTo", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GoTo", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx
		/// </summary>
		/// <param name="what">optional object What</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo(object what)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what);
			object returnItem = Invoker.MethodReturn(this, "GoTo", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx
		/// </summary>
		/// <param name="what">optional object What</param>
		/// <param name="which">optional object Which</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo(object what, object which)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, which);
			object returnItem = Invoker.MethodReturn(this, "GoTo", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx
		/// </summary>
		/// <param name="what">optional object What</param>
		/// <param name="which">optional object Which</param>
		/// <param name="count">optional object Count</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo(object what, object which, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, which, count);
			object returnItem = Invoker.MethodReturn(this, "GoTo", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836451.aspx
		/// </summary>
		/// <param name="what">NetOffice.WordApi.Enums.WdGoToItem What</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoToNext(NetOffice.WordApi.Enums.WdGoToItem what)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what);
			object returnItem = Invoker.MethodReturn(this, "GoToNext", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839107.aspx
		/// </summary>
		/// <param name="what">NetOffice.WordApi.Enums.WdGoToItem What</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoToPrevious(NetOffice.WordApi.Enums.WdGoToItem what)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what);
			object returnItem = Invoker.MethodReturn(this, "GoToPrevious", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="link">optional object Link</param>
		/// <param name="placement">optional object Placement</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="iconFileName">optional object IconFileName</param>
		/// <param name="iconLabel">optional object IconLabel</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType, object iconFileName, object iconLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex, link, placement, displayAsIcon, dataType, iconFileName, iconLabel);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="link">optional object Link</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex, object link)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex, link);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="link">optional object Link</param>
		/// <param name="placement">optional object Placement</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex, object link, object placement)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex, link, placement);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="link">optional object Link</param>
		/// <param name="placement">optional object Placement</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex, link, placement, displayAsIcon);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="link">optional object Link</param>
		/// <param name="placement">optional object Placement</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="dataType">optional object DataType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex, link, placement, displayAsIcon, dataType);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx
		/// </summary>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="link">optional object Link</param>
		/// <param name="placement">optional object Placement</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="iconFileName">optional object IconFileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType, object iconFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iconIndex, link, placement, displayAsIcon, dataType, iconFileName);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834516.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field PreviousField()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PreviousField", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194299.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field NextField()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NextField", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840515.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertParagraphBefore()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertParagraphBefore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194778.aspx
		/// </summary>
		/// <param name="shiftCells">optional object ShiftCells</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertCells(object shiftCells)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shiftCells);
			Invoker.Method(this, "InsertCells", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194778.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertCells()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertCells", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821815.aspx
		/// </summary>
		/// <param name="character">optional object Character</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Extend(object character)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(character);
			Invoker.Method(this, "Extend", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821815.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Extend()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Extend", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840081.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Shrink()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Shrink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		/// <param name="extend">optional object Extend</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveLeft(object unit, object count, object extend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count, extend);
			object returnItem = Invoker.MethodReturn(this, "MoveLeft", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveLeft()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "MoveLeft", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveLeft(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "MoveLeft", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveLeft(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "MoveLeft", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		/// <param name="extend">optional object Extend</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveRight(object unit, object count, object extend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count, extend);
			object returnItem = Invoker.MethodReturn(this, "MoveRight", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveRight()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "MoveRight", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveRight(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "MoveRight", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveRight(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "MoveRight", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		/// <param name="extend">optional object Extend</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveUp(object unit, object count, object extend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count, extend);
			object returnItem = Invoker.MethodReturn(this, "MoveUp", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveUp()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "MoveUp", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveUp(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "MoveUp", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveUp(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "MoveUp", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		/// <param name="extend">optional object Extend</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveDown(object unit, object count, object extend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count, extend);
			object returnItem = Invoker.MethodReturn(this, "MoveDown", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveDown()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "MoveDown", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveDown(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "MoveDown", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="count">optional object Count</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveDown(object unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "MoveDown", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192384.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="extend">optional object Extend</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 HomeKey(object unit, object extend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, extend);
			object returnItem = Invoker.MethodReturn(this, "HomeKey", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192384.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 HomeKey()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "HomeKey", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192384.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 HomeKey(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "HomeKey", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195593.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		/// <param name="extend">optional object Extend</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 EndKey(object unit, object extend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, extend);
			object returnItem = Invoker.MethodReturn(this, "EndKey", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195593.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 EndKey()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "EndKey", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195593.aspx
		/// </summary>
		/// <param name="unit">optional object Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 EndKey(object unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "EndKey", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835736.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void EscapeKey()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EscapeKey", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840867.aspx
		/// </summary>
		/// <param name="text">string Text</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void TypeText(string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			Invoker.Method(this, "TypeText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840230.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CopyFormat()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CopyFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196637.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PasteFormat()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PasteFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839799.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void TypeParagraph()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "TypeParagraph", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194909.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void TypeBackspace()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "TypeBackspace", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839790.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void NextSubdocument()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "NextSubdocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845750.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PreviousSubdocument()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PreviousSubdocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836022.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SelectColumn()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectColumn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197469.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SelectCurrentFont()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectCurrentFont", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822643.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SelectCurrentAlignment()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectCurrentAlignment", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191872.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SelectCurrentSpacing()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectCurrentSpacing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193883.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SelectCurrentIndent()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectCurrentIndent", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193718.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SelectCurrentTabs()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectCurrentTabs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840690.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SelectCurrentColor()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectCurrentColor", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839540.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateTextbox()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CreateTextbox", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840046.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void WholeStory()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "WholeStory", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845469.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SelectRow()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectRow", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196707.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SplitTable()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SplitTable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193340.aspx
		/// </summary>
		/// <param name="numRows">optional object NumRows</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertRows(object numRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numRows);
			Invoker.Method(this, "InsertRows", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193340.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertRows()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertRows", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838759.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertColumns()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertColumns", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835475.aspx
		/// </summary>
		/// <param name="formula">optional object Formula</param>
		/// <param name="numberFormat">optional object NumberFormat</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertFormula(object formula, object numberFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formula, numberFormat);
			Invoker.Method(this, "InsertFormula", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835475.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertFormula()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertFormula", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835475.aspx
		/// </summary>
		/// <param name="formula">optional object Formula</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertFormula(object formula)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formula);
			Invoker.Method(this, "InsertFormula", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834850.aspx
		/// </summary>
		/// <param name="wrap">optional object Wrap</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Revision NextRevision(object wrap)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wrap);
			object returnItem = Invoker.MethodReturn(this, "NextRevision", paramsArray);
			NetOffice.WordApi.Revision newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Revision.LateBindingApiWrapperType) as NetOffice.WordApi.Revision;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834850.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Revision NextRevision()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NextRevision", paramsArray);
			NetOffice.WordApi.Revision newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Revision.LateBindingApiWrapperType) as NetOffice.WordApi.Revision;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839603.aspx
		/// </summary>
		/// <param name="wrap">optional object Wrap</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Revision PreviousRevision(object wrap)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wrap);
			object returnItem = Invoker.MethodReturn(this, "PreviousRevision", paramsArray);
			NetOffice.WordApi.Revision newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Revision.LateBindingApiWrapperType) as NetOffice.WordApi.Revision;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839603.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Revision PreviousRevision()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PreviousRevision", paramsArray);
			NetOffice.WordApi.Revision newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Revision.LateBindingApiWrapperType) as NetOffice.WordApi.Revision;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194535.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PasteAsNestedTable()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PasteAsNestedTable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839331.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="styleName">string StyleName</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.AutoTextEntry CreateAutoTextEntry(string name, string styleName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, styleName);
			object returnItem = Invoker.MethodReturn(this, "CreateAutoTextEntry", paramsArray);
			NetOffice.WordApi.AutoTextEntry newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.AutoTextEntry.LateBindingApiWrapperType) as NetOffice.WordApi.AutoTextEntry;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838494.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DetectLanguage()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DetectLanguage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195143.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SelectCell()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectCell", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838718.aspx
		/// </summary>
		/// <param name="numRows">optional object NumRows</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertRowsBelow(object numRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numRows);
			Invoker.Method(this, "InsertRowsBelow", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838718.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertRowsBelow()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertRowsBelow", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844950.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertColumnsRight()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertColumnsRight", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840557.aspx
		/// </summary>
		/// <param name="numRows">optional object NumRows</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertRowsAbove(object numRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numRows);
			Invoker.Method(this, "InsertRowsAbove", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840557.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertRowsAbove()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertRowsAbove", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821034.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RtlRun()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RtlRun", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839502.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void LtrRun()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "LtrRun", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845275.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void BoldRun()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "BoldRun", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845442.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ItalicRun()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ItalicRun", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836904.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RtlPara()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RtlPara", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834853.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void LtrPara()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "LtrPara", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		/// <param name="insertAsField">optional object InsertAsField</param>
		/// <param name="insertAsFullWidth">optional object InsertAsFullWidth</param>
		/// <param name="dateLanguage">optional object DateLanguage</param>
		/// <param name="calendarType">optional object CalendarType</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth, object dateLanguage, object calendarType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField, insertAsFullWidth, dateLanguage, calendarType);
			Invoker.Method(this, "InsertDateTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertDateTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime(object dateTimeFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat);
			Invoker.Method(this, "InsertDateTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		/// <param name="insertAsField">optional object InsertAsField</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime(object dateTimeFormat, object insertAsField)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField);
			Invoker.Method(this, "InsertDateTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		/// <param name="insertAsField">optional object InsertAsField</param>
		/// <param name="insertAsFullWidth">optional object InsertAsFullWidth</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField, insertAsFullWidth);
			Invoker.Method(this, "InsertDateTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx
		/// </summary>
		/// <param name="dateTimeFormat">optional object DateTimeFormat</param>
		/// <param name="insertAsField">optional object InsertAsField</param>
		/// <param name="insertAsFullWidth">optional object InsertAsFullWidth</param>
		/// <param name="dateLanguage">optional object DateLanguage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth, object dateLanguage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField, insertAsFullWidth, dateLanguage);
			Invoker.Method(this, "InsertDateTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		/// <param name="subFieldNumber">optional object SubFieldNumber</param>
		/// <param name="subFieldNumber2">optional object SubFieldNumber2</param>
		/// <param name="subFieldNumber3">optional object SubFieldNumber3</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID, object subFieldNumber, object subFieldNumber2, object subFieldNumber3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID, subFieldNumber, subFieldNumber2, subFieldNumber3);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		/// <param name="subFieldNumber">optional object SubFieldNumber</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID, object subFieldNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID, subFieldNumber);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx
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
		/// <param name="subFieldNumber">optional object SubFieldNumber</param>
		/// <param name="subFieldNumber2">optional object SubFieldNumber2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID, object subFieldNumber, object subFieldNumber2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID, subFieldNumber, subFieldNumber2);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit, object autoFitBehavior, object defaultTableBehavior)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit, autoFitBehavior, defaultTableBehavior);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="numRows">optional object NumRows</param>
		/// <param name="numColumns">optional object NumColumns</param>
		/// <param name="initialColumnWidth">optional object InitialColumnWidth</param>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit, object autoFitBehavior)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit, autoFitBehavior);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTable", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe);
			Invoker.Method(this, "Sort2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197496.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ClearFormatting()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearFormatting", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196969.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PasteAppendTable()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PasteAppendTable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839633.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ToggleCharacterCode()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ToggleCharacterCode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821674.aspx
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdRecoveryType Type</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PasteAndFormat(NetOffice.WordApi.Enums.WdRecoveryType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "PasteAndFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837670.aspx
		/// </summary>
		/// <param name="linkedToExcel">bool LinkedToExcel</param>
		/// <param name="wordFormatting">bool WordFormatting</param>
		/// <param name="rTF">bool RTF</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PasteExcelTable(bool linkedToExcel, bool wordFormatting, bool rTF)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(linkedToExcel, wordFormatting, rTF);
			Invoker.Method(this, "PasteExcelTable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838352.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ShrinkDiscontiguousSelection()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ShrinkDiscontiguousSelection", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838293.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void InsertStyleSeparator()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertStyleSeparator", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		/// <param name="insertAsHyperlink">optional object InsertAsHyperlink</param>
		/// <param name="includePosition">optional object IncludePosition</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition);
			Invoker.Method(this, "InsertCrossReference_2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem);
			Invoker.Method(this, "InsertCrossReference_2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind ReferenceKind</param>
		/// <param name="referenceItem">object ReferenceItem</param>
		/// <param name="insertAsHyperlink">optional object InsertAsHyperlink</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType, referenceKind, referenceItem, insertAsHyperlink);
			Invoker.Method(this, "InsertCrossReference_2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		/// <param name="titleAutoText">optional object TitleAutoText</param>
		/// <param name="position">optional object Position</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void InsertCaptionXP(object label, object title, object titleAutoText, object position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title, titleAutoText, position);
			Invoker.Method(this, "InsertCaptionXP", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="label">object Label</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void InsertCaptionXP(object label)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label);
			Invoker.Method(this, "InsertCaptionXP", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void InsertCaptionXP(object label, object title)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title);
			Invoker.Method(this, "InsertCaptionXP", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="label">object Label</param>
		/// <param name="title">optional object Title</param>
		/// <param name="titleAutoText">optional object TitleAutoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void InsertCaptionXP(object label, object title, object titleAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(label, title, titleAutoText);
			Invoker.Method(this, "InsertCaptionXP", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844866.aspx
		/// </summary>
		/// <param name="editorID">optional object EditorID</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Range GoToEditableRange(object editorID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(editorID);
			object returnItem = Invoker.MethodReturn(this, "GoToEditableRange", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844866.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Range GoToEditableRange()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GoToEditableRange", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821369.aspx
		/// </summary>
		/// <param name="xML">string XML</param>
		/// <param name="transform">optional object Transform</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void InsertXML(string xML, object transform)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xML, transform);
			Invoker.Method(this, "InsertXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821369.aspx
		/// </summary>
		/// <param name="xML">string XML</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void InsertXML(string xML)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xML);
			Invoker.Method(this, "InsertXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838493.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ClearParagraphStyle()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearParagraphStyle", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191975.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ClearCharacterAllFormatting()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearCharacterAllFormatting", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841083.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ClearCharacterStyle()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearCharacterStyle", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838672.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ClearCharacterDirectFormatting()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearCharacterDirectFormatting", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx
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
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object fixedFormatExtClassPtr)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1, fixedFormatExtClassPtr);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx
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
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx
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
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx
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
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx
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
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx
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
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196419.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ReadingModeGrowFont()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ReadingModeGrowFont", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196279.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ReadingModeShrinkFont()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ReadingModeShrinkFont", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836876.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ClearParagraphAllFormatting()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearParagraphAllFormatting", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197502.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ClearParagraphDirectFormatting()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearParagraphDirectFormatting", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195985.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void InsertNewPage()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertNewPage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx
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
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SortByHeadings()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive, bidiSort);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx
		/// </summary>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		/// <param name="ignoreDiacritics">optional object IgnoreDiacritics</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx
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
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe);
			Invoker.Method(this, "SortByHeadings", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}