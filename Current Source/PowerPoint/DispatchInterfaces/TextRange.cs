using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using LateBindingApi.Core;
namespace NetOffice.PowerPointApi
{
	///<summary>
	/// DispatchInterface TextRange 
	/// SupportByVersion PowerPoint, 9,10,11,12,14
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class TextRange : Collection
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

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange(COMObject replacedObject) : base(replacedObject)
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PowerPointApi.Application newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Application.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.ActionSettings ActionSettings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActionSettings", paramsArray);
				NetOffice.PowerPointApi.ActionSettings newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.ActionSettings.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ActionSettings;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public Int32 Start
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Start", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Font Font
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Font", paramsArray);
				NetOffice.PowerPointApi.Font newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Font.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Font;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.ParagraphFormat ParagraphFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParagraphFormat", paramsArray);
				NetOffice.PowerPointApi.ParagraphFormat newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.ParagraphFormat.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ParagraphFormat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public Int32 IndentLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IndentLevel", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IndentLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="start">optional Int32 Start = -1</param>
		/// <param name="length">optional Int32 Length = -1</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Paragraphs(Int32 start, Int32 length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, length);
			object returnItem = Invoker.MethodReturn(this, "Paragraphs", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Paragraphs()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Paragraphs", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="start">optional Int32 Start = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Paragraphs(Int32 start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.MethodReturn(this, "Paragraphs", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="start">optional Int32 Start = -1</param>
		/// <param name="length">optional Int32 Length = -1</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Sentences(Int32 start, Int32 length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, length);
			object returnItem = Invoker.MethodReturn(this, "Sentences", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Sentences()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Sentences", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="start">optional Int32 Start = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Sentences(Int32 start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.MethodReturn(this, "Sentences", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="start">optional Int32 Start = -1</param>
		/// <param name="length">optional Int32 Length = -1</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Words(Int32 start, Int32 length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, length);
			object returnItem = Invoker.MethodReturn(this, "Words", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Words()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Words", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="start">optional Int32 Start = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Words(Int32 start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.MethodReturn(this, "Words", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="start">optional Int32 Start = -1</param>
		/// <param name="length">optional Int32 Length = -1</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Characters(Int32 start, Int32 length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, length);
			object returnItem = Invoker.MethodReturn(this, "Characters", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Characters()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Characters", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="start">optional Int32 Start = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Characters(Int32 start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.MethodReturn(this, "Characters", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="start">optional Int32 Start = -1</param>
		/// <param name="length">optional Int32 Length = -1</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Lines(Int32 start, Int32 length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, length);
			object returnItem = Invoker.MethodReturn(this, "Lines", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Lines()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Lines", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="start">optional Int32 Start = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Lines(Int32 start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.MethodReturn(this, "Lines", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="start">optional Int32 Start = -1</param>
		/// <param name="length">optional Int32 Length = -1</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Runs(Int32 start, Int32 length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, length);
			object returnItem = Invoker.MethodReturn(this, "Runs", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Runs()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Runs", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="start">optional Int32 Start = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Runs(Int32 start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.MethodReturn(this, "Runs", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange TrimText()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "TrimText", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="newText">optional string NewText = </param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange InsertAfter(string newText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(newText);
			object returnItem = Invoker.MethodReturn(this, "InsertAfter", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange InsertAfter()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "InsertAfter", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="newText">optional string NewText = </param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange InsertBefore(string newText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(newText);
			object returnItem = Invoker.MethodReturn(this, "InsertBefore", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange InsertBefore()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "InsertBefore", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="dateTimeFormat">NetOffice.PowerPointApi.Enums.PpDateTimeFormat DateTimeFormat</param>
		/// <param name="insertAsField">optional NetOffice.OfficeApi.Enums.MsoTriState InsertAsField = 0</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange InsertDateTime(NetOffice.PowerPointApi.Enums.PpDateTimeFormat dateTimeFormat, NetOffice.OfficeApi.Enums.MsoTriState insertAsField)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat, insertAsField);
			object returnItem = Invoker.MethodReturn(this, "InsertDateTime", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="dateTimeFormat">NetOffice.PowerPointApi.Enums.PpDateTimeFormat DateTimeFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange InsertDateTime(NetOffice.PowerPointApi.Enums.PpDateTimeFormat dateTimeFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateTimeFormat);
			object returnItem = Invoker.MethodReturn(this, "InsertDateTime", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange InsertSlideNumber()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "InsertSlideNumber", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fontName">string FontName</param>
		/// <param name="charNumber">Int32 CharNumber</param>
		/// <param name="unicode">optional NetOffice.OfficeApi.Enums.MsoTriState Unicode = 0</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange InsertSymbol(string fontName, Int32 charNumber, NetOffice.OfficeApi.Enums.MsoTriState unicode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fontName, charNumber, unicode);
			object returnItem = Invoker.MethodReturn(this, "InsertSymbol", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fontName">string FontName</param>
		/// <param name="charNumber">Int32 CharNumber</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange InsertSymbol(string fontName, Int32 charNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fontName, charNumber);
			object returnItem = Invoker.MethodReturn(this, "InsertSymbol", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Select()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Paste()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Paste", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpChangeCase Type</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void ChangeCase(NetOffice.PowerPointApi.Enums.PpChangeCase type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "ChangeCase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void AddPeriods()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddPeriods", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void RemovePeriods()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RemovePeriods", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="findWhat">string FindWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		/// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
		/// <param name="wholeWords">optional NetOffice.OfficeApi.Enums.MsoTriState WholeWords = 0</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Find(string findWhat, Int32 after, NetOffice.OfficeApi.Enums.MsoTriState matchCase, NetOffice.OfficeApi.Enums.MsoTriState wholeWords)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, after, matchCase, wholeWords);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="findWhat">string FindWhat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Find(string findWhat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="findWhat">string FindWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Find(string findWhat, Int32 after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, after);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="findWhat">string FindWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		/// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Find(string findWhat, Int32 after, NetOffice.OfficeApi.Enums.MsoTriState matchCase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, after, matchCase);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="findWhat">string FindWhat</param>
		/// <param name="replaceWhat">string ReplaceWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		/// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
		/// <param name="wholeWords">optional NetOffice.OfficeApi.Enums.MsoTriState WholeWords = 0</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Replace(string findWhat, string replaceWhat, Int32 after, NetOffice.OfficeApi.Enums.MsoTriState matchCase, NetOffice.OfficeApi.Enums.MsoTriState wholeWords)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, replaceWhat, after, matchCase, wholeWords);
			object returnItem = Invoker.MethodReturn(this, "Replace", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="findWhat">string FindWhat</param>
		/// <param name="replaceWhat">string ReplaceWhat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Replace(string findWhat, string replaceWhat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, replaceWhat);
			object returnItem = Invoker.MethodReturn(this, "Replace", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="findWhat">string FindWhat</param>
		/// <param name="replaceWhat">string ReplaceWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Replace(string findWhat, string replaceWhat, Int32 after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, replaceWhat, after);
			object returnItem = Invoker.MethodReturn(this, "Replace", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="findWhat">string FindWhat</param>
		/// <param name="replaceWhat">string ReplaceWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		/// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange Replace(string findWhat, string replaceWhat, Int32 after, NetOffice.OfficeApi.Enums.MsoTriState matchCase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, replaceWhat, after, matchCase);
			object returnItem = Invoker.MethodReturn(this, "Replace", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="x1">Single X1</param>
		/// <param name="y1">Single Y1</param>
		/// <param name="x2">Single X2</param>
		/// <param name="y2">Single Y2</param>
		/// <param name="x3">Single X3</param>
		/// <param name="y3">Single Y3</param>
		/// <param name="x4">Single x4</param>
		/// <param name="y4">Single y4</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void RtlRun()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RtlRun", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void LtrRun()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "LtrRun", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		/// <param name="iconLabel">optional string IconLabel = </param>
		/// <param name="link">optional NetOffice.OfficeApi.Enums.MsoTriState Link = 0</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial(NetOffice.PowerPointApi.Enums.PpPasteDataType dataType, NetOffice.OfficeApi.Enums.MsoTriState displayAsIcon, string iconFileName, Int32 iconIndex, string iconLabel, NetOffice.OfficeApi.Enums.MsoTriState link)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataType, displayAsIcon, iconFileName, iconIndex, iconLabel, link);
			object returnItem = Invoker.MethodReturn(this, "PasteSpecial", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PasteSpecial", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial(NetOffice.PowerPointApi.Enums.PpPasteDataType dataType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataType);
			object returnItem = Invoker.MethodReturn(this, "PasteSpecial", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial(NetOffice.PowerPointApi.Enums.PpPasteDataType dataType, NetOffice.OfficeApi.Enums.MsoTriState displayAsIcon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataType, displayAsIcon);
			object returnItem = Invoker.MethodReturn(this, "PasteSpecial", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial(NetOffice.PowerPointApi.Enums.PpPasteDataType dataType, NetOffice.OfficeApi.Enums.MsoTriState displayAsIcon, string iconFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataType, displayAsIcon, iconFileName);
			object returnItem = Invoker.MethodReturn(this, "PasteSpecial", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial(NetOffice.PowerPointApi.Enums.PpPasteDataType dataType, NetOffice.OfficeApi.Enums.MsoTriState displayAsIcon, string iconFileName, Int32 iconIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataType, displayAsIcon, iconFileName, iconIndex);
			object returnItem = Invoker.MethodReturn(this, "PasteSpecial", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		/// <param name="iconLabel">optional string IconLabel = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial(NetOffice.PowerPointApi.Enums.PpPasteDataType dataType, NetOffice.OfficeApi.Enums.MsoTriState displayAsIcon, string iconFileName, Int32 iconIndex, string iconLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataType, displayAsIcon, iconFileName, iconIndex, iconLabel);
			object returnItem = Invoker.MethodReturn(this, "PasteSpecial", paramsArray);
			NetOffice.PowerPointApi.TextRange newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextRange;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}