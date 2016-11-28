using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.PublisherApi
{
	///<summary>
	/// DispatchInterface TextRange 
	/// SupportByVersion Publisher, 14,15,16
	///</summary>
	[SupportByVersionAttribute("Publisher", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class TextRange : COMObject
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
                    _type = typeof(TextRange);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PublisherApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Application.LateBindingApiWrapperType) as NetOffice.PublisherApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Font Font
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Font", paramsArray);
				NetOffice.PublisherApi.Font newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Font.LateBindingApiWrapperType) as NetOffice.PublisherApi.Font;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Font", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 Length
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Length", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Single BoundLeft
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BoundLeft", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Single BoundHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BoundHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Single BoundTop
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BoundTop", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Single BoundWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BoundWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ParagraphFormat ParagraphFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParagraphFormat", paramsArray);
				NetOffice.PublisherApi.ParagraphFormat newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.ParagraphFormat.LateBindingApiWrapperType) as NetOffice.PublisherApi.ParagraphFormat;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ParagraphFormat", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public object ContainingObject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainingObject", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Duplicate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Duplicate", paramsArray);
				NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Font MajorityFont
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MajorityFont", paramsArray);
				NetOffice.PublisherApi.Font newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Font.LateBindingApiWrapperType) as NetOffice.PublisherApi.Font;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ParagraphFormat MajorityParagraphFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MajorityParagraphFormat", paramsArray);
				NetOffice.PublisherApi.ParagraphFormat newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.ParagraphFormat.LateBindingApiWrapperType) as NetOffice.PublisherApi.ParagraphFormat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Fields Fields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Fields", paramsArray);
				NetOffice.PublisherApi.Fields newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Fields.LateBindingApiWrapperType) as NetOffice.PublisherApi.Fields;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Story Story
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Story", paramsArray);
				NetOffice.PublisherApi.Story newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Story.LateBindingApiWrapperType) as NetOffice.PublisherApi.Story;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoLanguageID LanguageID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LanguageID", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoLanguageID)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LanguageID", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.DropCap DropCap
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DropCap", paramsArray);
				NetOffice.PublisherApi.DropCap newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.DropCap.LateBindingApiWrapperType) as NetOffice.PublisherApi.DropCap;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbFontScriptType Script
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Script", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PublisherApi.Enums.PbFontScriptType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlinks Hyperlinks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hyperlinks", paramsArray);
				NetOffice.PublisherApi.Hyperlinks newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Hyperlinks.LateBindingApiWrapperType) as NetOffice.PublisherApi.Hyperlinks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.FindReplace Find
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Find", paramsArray);
				NetOffice.PublisherApi.FindReplace newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.FindReplace.LateBindingApiWrapperType) as NetOffice.PublisherApi.FindReplace;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.InlineShapes InlineShapes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "InlineShapes", paramsArray);
				NetOffice.PublisherApi.InlineShapes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.InlineShapes.LateBindingApiWrapperType) as NetOffice.PublisherApi.InlineShapes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 WordsCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WordsCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 LinesCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LinesCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 ParagraphsCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParagraphsCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="direction">NetOffice.PublisherApi.Enums.PbCollapseDirection Direction</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Collapse(NetOffice.PublisherApi.Enums.PbCollapseDirection direction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(direction);
			Invoker.Method(this, "Collapse", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbTextUnit Unit</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 Expand(NetOffice.PublisherApi.Enums.PbTextUnit unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "Expand", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbTextUnit Unit</param>
		/// <param name="size">Int32 Size</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 Move(NetOffice.PublisherApi.Enums.PbTextUnit unit, Int32 size)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, size);
			object returnItem = Invoker.MethodReturn(this, "Move", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbTextUnit Unit</param>
		/// <param name="size">Int32 Size</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 MoveStart(NetOffice.PublisherApi.Enums.PbTextUnit unit, Int32 size)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, size);
			object returnItem = Invoker.MethodReturn(this, "MoveStart", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbTextUnit Unit</param>
		/// <param name="size">Int32 Size</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 MoveEnd(NetOffice.PublisherApi.Enums.PbTextUnit unit, Int32 size)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, size);
			object returnItem = Invoker.MethodReturn(this, "MoveEnd", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="start">Int32 Start</param>
		/// <param name="length">optional Int32 Length = 1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Characters(Int32 start, object length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, length);
			object returnItem = Invoker.MethodReturn(this, "Characters", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="start">Int32 Start</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Characters(Int32 start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.MethodReturn(this, "Characters", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="newText">string NewText</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertAfter(string newText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(newText);
			object returnItem = Invoker.MethodReturn(this, "InsertAfter", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="newText">string NewText</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertBefore(string newText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(newText);
			object returnItem = Invoker.MethodReturn(this, "InsertBefore", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fontName">string FontName</param>
		/// <param name="charIndex">Int32 CharIndex</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertSymbol(string fontName, Int32 charIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fontName, charIndex);
			object returnItem = Invoker.MethodReturn(this, "InsertSymbol", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat Format</param>
		/// <param name="insertAsField">optional bool InsertAsField = false</param>
		/// <param name="insertAsFullWidth">optional bool InsertAsFullWidth = false</param>
		/// <param name="language">optional NetOffice.OfficeApi.Enums.MsoLanguageID Language = 0</param>
		/// <param name="calendar">optional NetOffice.PublisherApi.Enums.PbCalendarType Calendar = 0</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format, object insertAsField, object insertAsFullWidth, object language, object calendar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, insertAsField, insertAsFullWidth, language, calendar);
			object returnItem = Invoker.MethodReturn(this, "InsertDateTime", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format);
			object returnItem = Invoker.MethodReturn(this, "InsertDateTime", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat Format</param>
		/// <param name="insertAsField">optional bool InsertAsField = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format, object insertAsField)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, insertAsField);
			object returnItem = Invoker.MethodReturn(this, "InsertDateTime", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat Format</param>
		/// <param name="insertAsField">optional bool InsertAsField = false</param>
		/// <param name="insertAsFullWidth">optional bool InsertAsFullWidth = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format, object insertAsField, object insertAsFullWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, insertAsField, insertAsFullWidth);
			object returnItem = Invoker.MethodReturn(this, "InsertDateTime", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat Format</param>
		/// <param name="insertAsField">optional bool InsertAsField = false</param>
		/// <param name="insertAsFullWidth">optional bool InsertAsFullWidth = false</param>
		/// <param name="language">optional NetOffice.OfficeApi.Enums.MsoLanguageID Language = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format, object insertAsField, object insertAsFullWidth, object language)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, insertAsField, insertAsFullWidth, language);
			object returnItem = Invoker.MethodReturn(this, "InsertDateTime", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="start">Int32 Start</param>
		/// <param name="length">optional Int32 Length = 1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Paragraphs(Int32 start, object length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, length);
			object returnItem = Invoker.MethodReturn(this, "Paragraphs", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="start">Int32 Start</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Paragraphs(Int32 start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.MethodReturn(this, "Paragraphs", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="start">Int32 Start</param>
		/// <param name="length">optional Int32 Length = 1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Lines(Int32 start, object length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, length);
			object returnItem = Invoker.MethodReturn(this, "Lines", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="start">Int32 Start</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Lines(Int32 start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.MethodReturn(this, "Lines", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="start">Int32 Start</param>
		/// <param name="length">optional Int32 Length = 1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Words(Int32 start, object length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, length);
			object returnItem = Invoker.MethodReturn(this, "Words", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="start">Int32 Start</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Words(Int32 start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.MethodReturn(this, "Words", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Select()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Paste()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Paste", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varIndex">object varIndex</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertMailMergeField(object varIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varIndex);
			object returnItem = Invoker.MethodReturn(this, "InsertMailMergeField", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.PublisherApi.Enums.PbPageNumberType Type = 1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertPageNumber(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "InsertPageNumber", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertPageNumber()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "InsertPageNumber", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertBarcode()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "InsertBarcode", paramsArray);
			NetOffice.PublisherApi.TextRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextRange;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}