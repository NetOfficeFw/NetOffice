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
	/// DispatchInterface _Document 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Document : COMObject
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
                    _type = typeof(_Document);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196900.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822944.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845529.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840549.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196862.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object BuiltInDocumentProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BuiltInDocumentProperties", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195603.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object CustomDocumentProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomDocumentProperties", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821867.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string Path
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Path", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194977.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835455.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197126.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194032.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845880.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823228.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdDocumentType Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdDocumentType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191749.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool AutoHyphenation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoHyphenation", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutoHyphenation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845783.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool HyphenateCaps
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HyphenateCaps", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HyphenateCaps", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193110.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 HyphenationZone
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HyphenationZone", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HyphenationZone", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820862.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 ConsecutiveHyphensLimit
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConsecutiveHyphensLimit", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ConsecutiveHyphensLimit", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822125.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836325.aspx
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
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845024.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194403.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191729.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821229.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840117.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193100.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Styles Styles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Styles", paramsArray);
				NetOffice.WordApi.Styles newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Styles.LateBindingApiWrapperType) as NetOffice.WordApi.Styles;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197117.aspx
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
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191950.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TablesOfFigures TablesOfFigures
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TablesOfFigures", paramsArray);
				NetOffice.WordApi.TablesOfFigures newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.TablesOfFigures.LateBindingApiWrapperType) as NetOffice.WordApi.TablesOfFigures;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834524.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Variables Variables
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Variables", paramsArray);
				NetOffice.WordApi.Variables newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Variables.LateBindingApiWrapperType) as NetOffice.WordApi.Variables;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198370.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMerge MailMerge
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailMerge", paramsArray);
				NetOffice.WordApi.MailMerge newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.MailMerge.LateBindingApiWrapperType) as NetOffice.WordApi.MailMerge;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844798.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Envelope Envelope
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Envelope", paramsArray);
				NetOffice.WordApi.Envelope newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Envelope.LateBindingApiWrapperType) as NetOffice.WordApi.Envelope;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821285.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string FullName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FullName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192540.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Revisions Revisions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Revisions", paramsArray);
				NetOffice.WordApi.Revisions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Revisions.LateBindingApiWrapperType) as NetOffice.WordApi.Revisions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822932.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TablesOfContents TablesOfContents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TablesOfContents", paramsArray);
				NetOffice.WordApi.TablesOfContents newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.TablesOfContents.LateBindingApiWrapperType) as NetOffice.WordApi.TablesOfContents;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837912.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TablesOfAuthorities TablesOfAuthorities
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TablesOfAuthorities", paramsArray);
				NetOffice.WordApi.TablesOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.TablesOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TablesOfAuthorities;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839306.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837336.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Windows Windows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Windows", paramsArray);
				NetOffice.WordApi.Windows newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Windows.LateBindingApiWrapperType) as NetOffice.WordApi.Windows;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool HasRoutingSlip
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasRoutingSlip", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HasRoutingSlip", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.RoutingSlip RoutingSlip
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RoutingSlip", paramsArray);
				NetOffice.WordApi.RoutingSlip newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.RoutingSlip.LateBindingApiWrapperType) as NetOffice.WordApi.RoutingSlip;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Routed
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Routed", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838095.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TablesOfAuthoritiesCategories TablesOfAuthoritiesCategories
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TablesOfAuthoritiesCategories", paramsArray);
				NetOffice.WordApi.TablesOfAuthoritiesCategories newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.TablesOfAuthoritiesCategories.LateBindingApiWrapperType) as NetOffice.WordApi.TablesOfAuthoritiesCategories;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194976.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Indexes Indexes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Indexes", paramsArray);
				NetOffice.WordApi.Indexes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Indexes.LateBindingApiWrapperType) as NetOffice.WordApi.Indexes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194753.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Saved
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Saved", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Saved", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821853.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Content
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Content", paramsArray);
				NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198228.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Window ActiveWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveWindow", paramsArray);
				NetOffice.WordApi.Window newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Window.LateBindingApiWrapperType) as NetOffice.WordApi.Window;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192728.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdDocumentKind Kind
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Kind", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdDocumentKind)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Kind", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196223.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool ReadOnly
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadOnly", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195362.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Subdocuments Subdocuments
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Subdocuments", paramsArray);
				NetOffice.WordApi.Subdocuments newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Subdocuments.LateBindingApiWrapperType) as NetOffice.WordApi.Subdocuments;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840840.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool IsMasterDocument
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsMasterDocument", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196079.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single DefaultTabStop
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultTabStop", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultTabStop", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836281.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool EmbedTrueTypeFonts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EmbedTrueTypeFonts", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EmbedTrueTypeFonts", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845567.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool SaveFormsData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SaveFormsData", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SaveFormsData", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838914.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool ReadOnlyRecommended
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadOnlyRecommended", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ReadOnlyRecommended", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844828.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool SaveSubsetFonts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SaveSubsetFonts", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SaveSubsetFonts", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840506.aspx
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdCompatibility Type</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_Compatibility(NetOffice.WordApi.Enums.WdCompatibility type)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.PropertyGet(this, "Compatibility", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdCompatibility Type</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Compatibility(NetOffice.WordApi.Enums.WdCompatibility type, bool value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.PropertySet(this, "Compatibility", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840506.aspx
		/// Alias for get_Compatibility
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdCompatibility Type</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Compatibility(NetOffice.WordApi.Enums.WdCompatibility type)
		{
			return get_Compatibility(type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197823.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.StoryRanges StoryRanges
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StoryRanges", paramsArray);
				NetOffice.WordApi.StoryRanges newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.StoryRanges.LateBindingApiWrapperType) as NetOffice.WordApi.StoryRanges;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821872.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CommandBars", paramsArray);
				NetOffice.OfficeApi.CommandBars newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192771.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool IsSubdocument
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsSubdocument", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840755.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 SaveFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SaveFormat", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836643.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdProtectionType ProtectionType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectionType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdProtectionType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837239.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197211.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shapes Shapes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shapes", paramsArray);
				NetOffice.WordApi.Shapes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Shapes.LateBindingApiWrapperType) as NetOffice.WordApi.Shapes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839163.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ListTemplates ListTemplates
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListTemplates", paramsArray);
				NetOffice.WordApi.ListTemplates newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ListTemplates.LateBindingApiWrapperType) as NetOffice.WordApi.ListTemplates;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845137.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Lists Lists
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Lists", paramsArray);
				NetOffice.WordApi.Lists newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Lists.LateBindingApiWrapperType) as NetOffice.WordApi.Lists;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821398.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool UpdateStylesOnOpen
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UpdateStylesOnOpen", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UpdateStylesOnOpen", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839734.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object AttachedTemplate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AttachedTemplate", paramsArray);
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
				Invoker.PropertySet(this, "AttachedTemplate", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844996.aspx
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
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844976.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape Background
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Background", paramsArray);
				NetOffice.WordApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Shape.LateBindingApiWrapperType) as NetOffice.WordApi.Shape;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Background", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193109.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845040.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836692.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool ShowGrammaticalErrors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowGrammaticalErrors", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowGrammaticalErrors", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821056.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool ShowSpellingErrors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowSpellingErrors", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowSpellingErrors", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Versions Versions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Versions", paramsArray);
				NetOffice.WordApi.Versions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Versions.LateBindingApiWrapperType) as NetOffice.WordApi.Versions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool ShowSummary
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowSummary", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowSummary", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSummaryMode SummaryViewMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SummaryViewMode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdSummaryMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SummaryViewMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 SummaryLength
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SummaryLength", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SummaryLength", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool PrintFractionalWidths
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintFractionalWidths", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintFractionalWidths", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196987.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool PrintPostScriptOverText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintPostScriptOverText", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintPostScriptOverText", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840423.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Container
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Container", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838735.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool PrintFormsData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintFormsData", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintFormsData", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198090.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ListParagraphs ListParagraphs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListParagraphs", paramsArray);
				NetOffice.WordApi.ListParagraphs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ListParagraphs.LateBindingApiWrapperType) as NetOffice.WordApi.ListParagraphs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192387.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string Password
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Password", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Password", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839518.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string WritePassword
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WritePassword", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "WritePassword", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194500.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool HasPassword
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasPassword", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837527.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool WriteReserved
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WriteReserved", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844946.aspx
		/// </summary>
		/// <param name="languageID">object LanguageID</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_ActiveWritingStyle(object languageID)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(languageID);
			object returnItem = Invoker.PropertyGet(this, "ActiveWritingStyle", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="languageID">object LanguageID</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ActiveWritingStyle(object languageID, string value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(languageID);
			Invoker.PropertySet(this, "ActiveWritingStyle", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844946.aspx
		/// Alias for get_ActiveWritingStyle
		/// </summary>
		/// <param name="languageID">object LanguageID</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string ActiveWritingStyle(object languageID)
		{
			return get_ActiveWritingStyle(languageID);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193401.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool UserControl
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UserControl", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UserControl", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool HasMailer
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasMailer", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HasMailer", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Mailer Mailer
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Mailer", paramsArray);
				NetOffice.WordApi.Mailer newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Mailer.LateBindingApiWrapperType) as NetOffice.WordApi.Mailer;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839868.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ReadabilityStatistics ReadabilityStatistics
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadabilityStatistics", paramsArray);
				NetOffice.WordApi.ReadabilityStatistics newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ReadabilityStatistics.LateBindingApiWrapperType) as NetOffice.WordApi.ReadabilityStatistics;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192400.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ProofreadingErrors GrammaticalErrors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GrammaticalErrors", paramsArray);
				NetOffice.WordApi.ProofreadingErrors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ProofreadingErrors.LateBindingApiWrapperType) as NetOffice.WordApi.ProofreadingErrors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838118.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ProofreadingErrors SpellingErrors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SpellingErrors", paramsArray);
				NetOffice.WordApi.ProofreadingErrors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ProofreadingErrors.LateBindingApiWrapperType) as NetOffice.WordApi.ProofreadingErrors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837668.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.VBIDEApi.VBProject VBProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBProject", paramsArray);
				NetOffice.VBIDEApi.VBProject newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.VBProject.LateBindingApiWrapperType) as NetOffice.VBIDEApi.VBProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840586.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool FormsDesign
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormsDesign", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string _CodeName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "_CodeName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "_CodeName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197577.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string CodeName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CodeName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821373.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool SnapToGrid
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SnapToGrid", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SnapToGrid", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837193.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool SnapToShapes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SnapToShapes", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SnapToShapes", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839124.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single GridDistanceHorizontal
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GridDistanceHorizontal", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GridDistanceHorizontal", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195287.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single GridDistanceVertical
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GridDistanceVertical", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GridDistanceVertical", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839558.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single GridOriginHorizontal
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GridOriginHorizontal", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GridOriginHorizontal", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198193.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single GridOriginVertical
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GridOriginVertical", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GridOriginVertical", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821306.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 GridSpaceBetweenHorizontalLines
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GridSpaceBetweenHorizontalLines", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GridSpaceBetweenHorizontalLines", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821996.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 GridSpaceBetweenVerticalLines
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GridSpaceBetweenVerticalLines", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GridSpaceBetweenVerticalLines", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845752.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool GridOriginFromMargin
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GridOriginFromMargin", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GridOriginFromMargin", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836931.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool KerningByAlgorithm
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "KerningByAlgorithm", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "KerningByAlgorithm", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191748.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdJustificationMode JustificationMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "JustificationMode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdJustificationMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "JustificationMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845667.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdFarEastLineBreakLevel FarEastLineBreakLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FarEastLineBreakLevel", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdFarEastLineBreakLevel)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FarEastLineBreakLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844966.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string NoLineBreakBefore
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NoLineBreakBefore", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "NoLineBreakBefore", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192597.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string NoLineBreakAfter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NoLineBreakAfter", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "NoLineBreakAfter", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838067.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool TrackRevisions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TrackRevisions", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TrackRevisions", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192825.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool PrintRevisions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintRevisions", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintRevisions", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool ShowRevisions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowRevisions", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowRevisions", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192741.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string ActiveTheme
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveTheme", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837037.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string ActiveThemeDisplayName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveThemeDisplayName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839292.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Email Email
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Email", paramsArray);
				NetOffice.WordApi.Email newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Email.LateBindingApiWrapperType) as NetOffice.WordApi.Email;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196093.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Scripts Scripts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Scripts", paramsArray);
				NetOffice.OfficeApi.Scripts newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Scripts.LateBindingApiWrapperType) as NetOffice.OfficeApi.Scripts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191794.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838486.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdFarEastLineBreakLanguageID FarEastLineBreakLanguage
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FarEastLineBreakLanguage", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdFarEastLineBreakLanguageID)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FarEastLineBreakLanguage", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194305.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Frameset Frameset
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Frameset", paramsArray);
				NetOffice.WordApi.Frameset newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Frameset.LateBindingApiWrapperType) as NetOffice.WordApi.Frameset;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839615.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ClickAndTypeParagraphStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ClickAndTypeParagraphStyle", paramsArray);
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
				Invoker.PropertySet(this, "ClickAndTypeParagraphStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.HTMLProject HTMLProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HTMLProject", paramsArray);
				NetOffice.OfficeApi.HTMLProject newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.HTMLProject.LateBindingApiWrapperType) as NetOffice.OfficeApi.HTMLProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844954.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.WebOptions WebOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WebOptions", paramsArray);
				NetOffice.WordApi.WebOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.WebOptions.LateBindingApiWrapperType) as NetOffice.WordApi.WebOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835467.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoEncoding OpenEncoding
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OpenEncoding", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoEncoding)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834893.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoEncoding SaveEncoding
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SaveEncoding", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoEncoding)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SaveEncoding", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835162.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool OptimizeForWord97
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OptimizeForWord97", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OptimizeForWord97", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836069.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool VBASigned
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBASigned", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840465.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.MsoEnvelope MailEnvelope
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailEnvelope", paramsArray);
				NetOffice.OfficeApi.MsoEnvelope newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.MsoEnvelope.LateBindingApiWrapperType) as NetOffice.OfficeApi.MsoEnvelope;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194348.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool DisableFeatures
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisableFeatures", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisableFeatures", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194604.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool DoNotEmbedSystemFonts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DoNotEmbedSystemFonts", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DoNotEmbedSystemFonts", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193069.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.SignatureSet Signatures
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Signatures", paramsArray);
				NetOffice.OfficeApi.SignatureSet newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SignatureSet.LateBindingApiWrapperType) as NetOffice.OfficeApi.SignatureSet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194661.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public string DefaultTargetFrame
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultTargetFrame", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultTargetFrame", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822985.aspx
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
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196211.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdDisableFeaturesIntroducedAfter DisableFeaturesIntroducedAfter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisableFeaturesIntroducedAfter", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdDisableFeaturesIntroducedAfter)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisableFeaturesIntroducedAfter", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838361.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool RemovePersonalInformation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RemovePersonalInformation", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RemovePersonalInformation", paramsArray);
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
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool EmbedSmartTags
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EmbedSmartTags", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EmbedSmartTags", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool SmartTagsAsXMLProps
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartTagsAsXMLProps", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SmartTagsAsXMLProps", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835460.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoEncoding TextEncoding
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextEncoding", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoEncoding)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TextEncoding", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198078.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdLineEndingType TextLineEnding
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextLineEnding", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdLineEndingType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TextLineEnding", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845673.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.StyleSheets StyleSheets
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StyleSheets", paramsArray);
				NetOffice.WordApi.StyleSheets newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.StyleSheets.LateBindingApiWrapperType) as NetOffice.WordApi.StyleSheets;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837042.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public object DefaultTableStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultTableStyle", paramsArray);
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194870.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public string PasswordEncryptionProvider
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PasswordEncryptionProvider", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195788.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public string PasswordEncryptionAlgorithm
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PasswordEncryptionAlgorithm", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193119.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public Int32 PasswordEncryptionKeyLength
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PasswordEncryptionKeyLength", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822966.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool PasswordEncryptionFileProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PasswordEncryptionFileProperties", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836336.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool EmbedLinguisticData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EmbedLinguisticData", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EmbedLinguisticData", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839893.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool FormattingShowFont
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormattingShowFont", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormattingShowFont", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839706.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool FormattingShowClear
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormattingShowClear", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormattingShowClear", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836749.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool FormattingShowParagraph
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormattingShowParagraph", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormattingShowParagraph", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193041.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool FormattingShowNumbering
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormattingShowNumbering", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormattingShowNumbering", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194361.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdShowFilter FormattingShowFilter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormattingShowFilter", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdShowFilter)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormattingShowFilter", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191744.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Permission Permission
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Permission", paramsArray);
				NetOffice.OfficeApi.Permission newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Permission.LateBindingApiWrapperType) as NetOffice.OfficeApi.Permission;
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198201.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLSchemaReferences XMLSchemaReferences
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLSchemaReferences", paramsArray);
				NetOffice.WordApi.XMLSchemaReferences newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.XMLSchemaReferences.LateBindingApiWrapperType) as NetOffice.WordApi.XMLSchemaReferences;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840776.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SmartDocument SmartDocument
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartDocument", paramsArray);
				NetOffice.OfficeApi.SmartDocument newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartDocument.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartDocument;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspace SharedWorkspace
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SharedWorkspace", paramsArray);
				NetOffice.OfficeApi.SharedWorkspace newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SharedWorkspace.LateBindingApiWrapperType) as NetOffice.OfficeApi.SharedWorkspace;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837910.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Sync Sync
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Sync", paramsArray);
				NetOffice.OfficeApi.Sync newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Sync.LateBindingApiWrapperType) as NetOffice.OfficeApi.Sync;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838344.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public bool EnforceStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnforceStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnforceStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822185.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public bool AutoFormatOverride
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoFormatOverride", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutoFormatOverride", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public bool XMLSaveDataOnly
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLSaveDataOnly", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "XMLSaveDataOnly", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public bool XMLHideNamespaces
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLHideNamespaces", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "XMLHideNamespaces", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196205.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public bool XMLShowAdvancedErrors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLShowAdvancedErrors", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "XMLShowAdvancedErrors", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836689.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public bool XMLUseXSLTWhenSaving
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLUseXSLTWhenSaving", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "XMLUseXSLTWhenSaving", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838300.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public string XMLSaveThroughXSLT
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLSaveThroughXSLT", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "XMLSaveThroughXSLT", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191946.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentLibraryVersions", paramsArray);
				NetOffice.OfficeApi.DocumentLibraryVersions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.DocumentLibraryVersions.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentLibraryVersions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196654.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public bool ReadingModeLayoutFrozen
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadingModeLayoutFrozen", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ReadingModeLayoutFrozen", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194610.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public bool RemoveDateAndTime
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RemoveDateAndTime", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RemoveDateAndTime", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLChildNodeSuggestions ChildNodeSuggestions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChildNodeSuggestions", paramsArray);
				NetOffice.WordApi.XMLChildNodeSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.XMLChildNodeSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.XMLChildNodeSuggestions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes XMLSchemaViolations
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLSchemaViolations", paramsArray);
				NetOffice.WordApi.XMLNodes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType) as NetOffice.WordApi.XMLNodes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191938.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public Int32 ReadingLayoutSizeX
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadingLayoutSizeX", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ReadingLayoutSizeX", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839167.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public Int32 ReadingLayoutSizeY
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadingLayoutSizeY", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ReadingLayoutSizeY", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191767.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdStyleSort StyleSortMethod
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StyleSortMethod", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdStyleSort)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "StyleSortMethod", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844919.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.MetaProperties ContentTypeProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContentTypeProperties", paramsArray);
				NetOffice.OfficeApi.MetaProperties newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.MetaProperties.LateBindingApiWrapperType) as NetOffice.OfficeApi.MetaProperties;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197907.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool TrackMoves
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TrackMoves", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TrackMoves", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836881.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool TrackFormatting
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TrackFormatting", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TrackFormatting", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Dummy1
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dummy1", paramsArray);
				return (object)returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837488.aspx
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
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Dummy3
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dummy3", paramsArray);
				return (object)returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839289.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.ServerPolicy ServerPolicy
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ServerPolicy", paramsArray);
				NetOffice.OfficeApi.ServerPolicy newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.ServerPolicy.LateBindingApiWrapperType) as NetOffice.OfficeApi.ServerPolicy;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822382.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839144.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.DocumentInspectors DocumentInspectors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentInspectors", paramsArray);
				NetOffice.OfficeApi.DocumentInspectors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.DocumentInspectors.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentInspectors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834552.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Bibliography Bibliography
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Bibliography", paramsArray);
				NetOffice.WordApi.Bibliography newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Bibliography.LateBindingApiWrapperType) as NetOffice.WordApi.Bibliography;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198209.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool LockTheme
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LockTheme", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LockTheme", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839340.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool LockQuickStyleSet
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LockQuickStyleSet", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LockQuickStyleSet", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821063.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public string OriginalDocumentTitle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OriginalDocumentTitle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834817.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public string RevisedDocumentTitle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RevisedDocumentTitle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193091.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLParts CustomXMLParts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomXMLParts", paramsArray);
				NetOffice.OfficeApi.CustomXMLParts newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLParts.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLParts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195284.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool FormattingShowNextLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormattingShowNextLevel", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormattingShowNextLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191723.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool FormattingShowUserStyleName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormattingShowUserStyleName", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormattingShowUserStyleName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822952.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Research Research
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Research", paramsArray);
				NetOffice.WordApi.Research newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Research.LateBindingApiWrapperType) as NetOffice.WordApi.Research;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838930.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool Final
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Final", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Final", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821662.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathBreakBin OMathBreakBin
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMathBreakBin", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdOMathBreakBin)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OMathBreakBin", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835681.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathBreakSub OMathBreakSub
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMathBreakSub", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdOMathBreakSub)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OMathBreakSub", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196528.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathJc OMathJc
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMathJc", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdOMathJc)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OMathJc", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195080.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public Single OMathLeftMargin
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMathLeftMargin", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OMathLeftMargin", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192826.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public Single OMathRightMargin
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMathRightMargin", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OMathRightMargin", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195018.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public Single OMathWrap
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMathWrap", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OMathWrap", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822912.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool OMathIntSubSupLim
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMathIntSubSupLim", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OMathIntSubSupLim", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192808.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool OMathNarySupSubLim
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMathNarySupSubLim", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OMathNarySupSubLim", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835679.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool OMathSmallFrac
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMathSmallFrac", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OMathSmallFrac", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197690.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840566.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.OfficeTheme DocumentTheme
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentTheme", paramsArray);
				NetOffice.OfficeApi.OfficeTheme newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.OfficeTheme.LateBindingApiWrapperType) as NetOffice.OfficeApi.OfficeTheme;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845747.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool HasVBProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasVBProject", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193851.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public string OMathFontName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMathFontName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OMathFontName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836379.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public string EncryptionProvider
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EncryptionProvider", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EncryptionProvider", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838359.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool UseMathDefaults
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UseMathDefaults", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UseMathDefaults", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195620.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public Int32 CurrentRsid
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentRsid", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 DocID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocID", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196837.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 CompatibilityMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CompatibilityMode", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837045.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.CoAuthoring CoAuthoring
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CoAuthoring", paramsArray);
				NetOffice.WordApi.CoAuthoring newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.CoAuthoring.LateBindingApiWrapperType) as NetOffice.WordApi.CoAuthoring;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231858.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 15, 16)]
		public NetOffice.WordApi.Broadcast Broadcast
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Broadcast", paramsArray);
				NetOffice.WordApi.Broadcast newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Broadcast.LateBindingApiWrapperType) as NetOffice.WordApi.Broadcast;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj228844.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 15, 16)]
		public bool ChartDataPointTrack
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartDataPointTrack", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ChartDataPointTrack", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230857.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 15, 16)]
		public bool IsInAutosave
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsInAutosave", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196343.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		/// <param name="originalFormat">optional object OriginalFormat</param>
		/// <param name="routeDocument">optional object RouteDocument</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges, object originalFormat, object routeDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, originalFormat, routeDocument);
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196343.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196343.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196343.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		/// <param name="originalFormat">optional object OriginalFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges, object originalFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, originalFormat);
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object AllowSubstitutions</param>
		/// <param name="lineEnding">optional object LineEnding</param>
		/// <param name="addBiDiMarks">optional object AddBiDiMarks</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SaveAs()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object AllowSubstitutions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object AllowSubstitutions</param>
		/// <param name="lineEnding">optional object LineEnding</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821326.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Repaginate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Repaginate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822617.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void FitToPages()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FitToPages", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841098.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ManualHyphenation()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ManualHyphenation", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845112.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Select()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845755.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DataForm()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DataForm", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Route()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Route", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821625.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Save()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821630.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SendMail()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendMail", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821608.aspx
		/// </summary>
		/// <param name="start">optional object Start</param>
		/// <param name="end">optional object End</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Range(object start, object end)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, end);
			object returnItem = Invoker.MethodReturn(this, "Range", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821608.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Range()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Range", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821608.aspx
		/// </summary>
		/// <param name="start">optional object Start</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Range(object start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.MethodReturn(this, "Range", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823210.aspx
		/// </summary>
		/// <param name="which">NetOffice.WordApi.Enums.WdAutoMacros Which</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RunAutoMacro(NetOffice.WordApi.Enums.WdAutoMacros which)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(which);
			Invoker.Method(this, "RunAutoMacro", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822131.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Activate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195898.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintPreview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintPreview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840796.aspx
		/// </summary>
		/// <param name="times">optional object Times</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Undo(object times)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(times);
			object returnItem = Invoker.MethodReturn(this, "Undo", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840796.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Undo()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Undo", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845577.aspx
		/// </summary>
		/// <param name="times">optional object Times</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Redo(object times)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(times);
			object returnItem = Invoker.MethodReturn(this, "Redo", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845577.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Redo()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Redo", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840638.aspx
		/// </summary>
		/// <param name="statistic">NetOffice.WordApi.Enums.WdStatistic Statistic</param>
		/// <param name="includeFootnotesAndEndnotes">optional object IncludeFootnotesAndEndnotes</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 ComputeStatistics(NetOffice.WordApi.Enums.WdStatistic statistic, object includeFootnotesAndEndnotes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(statistic, includeFootnotesAndEndnotes);
			object returnItem = Invoker.MethodReturn(this, "ComputeStatistics", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840638.aspx
		/// </summary>
		/// <param name="statistic">NetOffice.WordApi.Enums.WdStatistic Statistic</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 ComputeStatistics(NetOffice.WordApi.Enums.WdStatistic statistic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(statistic);
			object returnItem = Invoker.MethodReturn(this, "ComputeStatistics", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845133.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MakeCompatibilityDefault()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MakeCompatibilityDefault", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType Type</param>
		/// <param name="noReset">optional object NoReset</param>
		/// <param name="password">optional object Password</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, noReset, password);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType Type</param>
		/// <param name="noReset">optional object NoReset</param>
		/// <param name="password">optional object Password</param>
		/// <param name="useIRM">optional object UseIRM</param>
		/// <param name="enforceStyleLock">optional object EnforceStyleLock</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password, object useIRM, object enforceStyleLock)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, noReset, password, useIRM, enforceStyleLock);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Protect(NetOffice.WordApi.Enums.WdProtectionType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType Type</param>
		/// <param name="noReset">optional object NoReset</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, noReset);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType Type</param>
		/// <param name="noReset">optional object NoReset</param>
		/// <param name="password">optional object Password</param>
		/// <param name="useIRM">optional object UseIRM</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password, object useIRM)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, noReset, password, useIRM);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845016.aspx
		/// </summary>
		/// <param name="password">optional object Password</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Unprotect(object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password);
			Invoker.Method(this, "Unprotect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845016.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Unprotect()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Unprotect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdEditionType Type</param>
		/// <param name="option">NetOffice.WordApi.Enums.WdEditionOption Option</param>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void EditionOptions(NetOffice.WordApi.Enums.WdEditionType type, NetOffice.WordApi.Enums.WdEditionOption option, string name, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, option, name, format);
			Invoker.Method(this, "EditionOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdEditionType Type</param>
		/// <param name="option">NetOffice.WordApi.Enums.WdEditionOption Option</param>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void EditionOptions(NetOffice.WordApi.Enums.WdEditionType type, NetOffice.WordApi.Enums.WdEditionOption option, string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, option, name);
			Invoker.Method(this, "EditionOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821600.aspx
		/// </summary>
		/// <param name="letterContent">optional object LetterContent</param>
		/// <param name="wizardMode">optional object WizardMode</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RunLetterWizard(object letterContent, object wizardMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(letterContent, wizardMode);
			Invoker.Method(this, "RunLetterWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821600.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RunLetterWizard()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RunLetterWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821600.aspx
		/// </summary>
		/// <param name="letterContent">optional object LetterContent</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RunLetterWizard(object letterContent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(letterContent);
			Invoker.Method(this, "RunLetterWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836106.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent GetLetterContent()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetLetterContent", paramsArray);
			NetOffice.WordApi.LetterContent newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.LetterContent.LateBindingApiWrapperType) as NetOffice.WordApi.LetterContent;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822930.aspx
		/// </summary>
		/// <param name="letterContent">object LetterContent</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SetLetterContent(object letterContent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(letterContent);
			Invoker.Method(this, "SetLetterContent", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840260.aspx
		/// </summary>
		/// <param name="template">string Template</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CopyStylesFromTemplate(string template)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(template);
			Invoker.Method(this, "CopyStylesFromTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840983.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void UpdateStyles()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UpdateStyles", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834835.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckGrammar()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckGrammar", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx
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
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx
		/// </summary>
		/// <param name="address">optional object Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="newWindow">optional object NewWindow</param>
		/// <param name="addHistory">optional object AddHistory</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		/// <param name="method">optional object Method</param>
		/// <param name="headerInfo">optional object HeaderInfo</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx
		/// </summary>
		/// <param name="address">optional object Address</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx
		/// </summary>
		/// <param name="address">optional object Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address, object subAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx
		/// </summary>
		/// <param name="address">optional object Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="newWindow">optional object NewWindow</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address, object subAddress, object newWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx
		/// </summary>
		/// <param name="address">optional object Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="newWindow">optional object NewWindow</param>
		/// <param name="addHistory">optional object AddHistory</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx
		/// </summary>
		/// <param name="address">optional object Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="newWindow">optional object NewWindow</param>
		/// <param name="addHistory">optional object AddHistory</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory, object extraInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx
		/// </summary>
		/// <param name="address">optional object Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="newWindow">optional object NewWindow</param>
		/// <param name="addHistory">optional object AddHistory</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		/// <param name="method">optional object Method</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory, object extraInfo, object method)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo, method);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839781.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AddToFavorites()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddToFavorites", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195614.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Reload()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Reload", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="length">optional object Length</param>
		/// <param name="mode">optional object Mode</param>
		/// <param name="updateProperties">optional object UpdateProperties</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range AutoSummarize(object length, object mode, object updateProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(length, mode, updateProperties);
			object returnItem = Invoker.MethodReturn(this, "AutoSummarize", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range AutoSummarize()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AutoSummarize", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="length">optional object Length</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range AutoSummarize(object length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(length);
			object returnItem = Invoker.MethodReturn(this, "AutoSummarize", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="length">optional object Length</param>
		/// <param name="mode">optional object Mode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range AutoSummarize(object length, object mode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(length, mode);
			object returnItem = Invoker.MethodReturn(this, "AutoSummarize", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193060.aspx
		/// </summary>
		/// <param name="numberType">optional object NumberType</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RemoveNumbers(object numberType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numberType);
			Invoker.Method(this, "RemoveNumbers", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193060.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RemoveNumbers()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RemoveNumbers", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838874.aspx
		/// </summary>
		/// <param name="numberType">optional object NumberType</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ConvertNumbersToText(object numberType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numberType);
			Invoker.Method(this, "ConvertNumbersToText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838874.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ConvertNumbersToText()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ConvertNumbersToText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836858.aspx
		/// </summary>
		/// <param name="numberType">optional object NumberType</param>
		/// <param name="level">optional object Level</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems(object numberType, object level)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numberType, level);
			object returnItem = Invoker.MethodReturn(this, "CountNumberedItems", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836858.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CountNumberedItems", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836858.aspx
		/// </summary>
		/// <param name="numberType">optional object NumberType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems(object numberType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numberType);
			object returnItem = Invoker.MethodReturn(this, "CountNumberedItems", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192151.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Post()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Post", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195394.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ToggleFormsDesign()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ToggleFormsDesign", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Compare(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "Compare", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="authorName">optional object AuthorName</param>
		/// <param name="compareTarget">optional object CompareTarget</param>
		/// <param name="detectFormatChanges">optional object DetectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object IgnoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings, addToRecentFiles);
			Invoker.Method(this, "Compare", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="authorName">optional object AuthorName</param>
		/// <param name="compareTarget">optional object CompareTarget</param>
		/// <param name="detectFormatChanges">optional object DetectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object IgnoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="removePersonalInformation">optional object RemovePersonalInformation</param>
		/// <param name="removeDateAndTime">optional object RemoveDateAndTime</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles, object removePersonalInformation, object removeDateAndTime)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings, addToRecentFiles, removePersonalInformation, removeDateAndTime);
			Invoker.Method(this, "Compare", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="authorName">optional object AuthorName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Compare(string name, object authorName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, authorName);
			Invoker.Method(this, "Compare", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="authorName">optional object AuthorName</param>
		/// <param name="compareTarget">optional object CompareTarget</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Compare(string name, object authorName, object compareTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, authorName, compareTarget);
			Invoker.Method(this, "Compare", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="authorName">optional object AuthorName</param>
		/// <param name="compareTarget">optional object CompareTarget</param>
		/// <param name="detectFormatChanges">optional object DetectFormatChanges</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Compare(string name, object authorName, object compareTarget, object detectFormatChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, authorName, compareTarget, detectFormatChanges);
			Invoker.Method(this, "Compare", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="authorName">optional object AuthorName</param>
		/// <param name="compareTarget">optional object CompareTarget</param>
		/// <param name="detectFormatChanges">optional object DetectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object IgnoreAllComparisonWarnings</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings);
			Invoker.Method(this, "Compare", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="authorName">optional object AuthorName</param>
		/// <param name="compareTarget">optional object CompareTarget</param>
		/// <param name="detectFormatChanges">optional object DetectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object IgnoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="removePersonalInformation">optional object RemovePersonalInformation</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles, object removePersonalInformation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings, addToRecentFiles, removePersonalInformation);
			Invoker.Method(this, "Compare", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void UpdateSummaryProperties()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UpdateSummaryProperties", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193699.aspx
		/// </summary>
		/// <param name="referenceType">object ReferenceType</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object GetCrossReferenceItems(object referenceType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(referenceType);
			object returnItem = Invoker.MethodReturn(this, "GetCrossReferenceItems", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193992.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837880.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ViewCode()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ViewCode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834519.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ViewPropertyBrowser()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ViewPropertyBrowser", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ForwardMailer()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ForwardMailer", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Reply()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Reply", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ReplyAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ReplyAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="priority">optional object Priority</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SendMailer(object fileFormat, object priority)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileFormat, priority);
			Invoker.Method(this, "SendMailer", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SendMailer()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendMailer", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileFormat">optional object FileFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SendMailer(object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileFormat);
			Invoker.Method(this, "SendMailer", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195616.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void UndoClear()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UndoClear", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192417.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PresentIt()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PresentIt", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838927.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subject">optional object Subject</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SendFax(string address, object subject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subject);
			Invoker.Method(this, "SendFax", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838927.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SendFax(string address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address);
			Invoker.Method(this, "SendFax", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Merge(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "Merge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="mergeTarget">optional object MergeTarget</param>
		/// <param name="detectFormatChanges">optional object DetectFormatChanges</param>
		/// <param name="useFormattingFrom">optional object UseFormattingFrom</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Merge(string fileName, object mergeTarget, object detectFormatChanges, object useFormattingFrom, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, mergeTarget, detectFormatChanges, useFormattingFrom, addToRecentFiles);
			Invoker.Method(this, "Merge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="mergeTarget">optional object MergeTarget</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Merge(string fileName, object mergeTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, mergeTarget);
			Invoker.Method(this, "Merge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="mergeTarget">optional object MergeTarget</param>
		/// <param name="detectFormatChanges">optional object DetectFormatChanges</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Merge(string fileName, object mergeTarget, object detectFormatChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, mergeTarget, detectFormatChanges);
			Invoker.Method(this, "Merge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="mergeTarget">optional object MergeTarget</param>
		/// <param name="detectFormatChanges">optional object DetectFormatChanges</param>
		/// <param name="useFormattingFrom">optional object UseFormattingFrom</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Merge(string fileName, object mergeTarget, object detectFormatChanges, object useFormattingFrom)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, mergeTarget, detectFormatChanges, useFormattingFrom);
			Invoker.Method(this, "Merge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822702.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ClosePrintPreview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClosePrintPreview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834920.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CheckConsistency()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckConsistency", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx
		/// </summary>
		/// <param name="dateFormat">string DateFormat</param>
		/// <param name="includeHeaderFooter">bool IncludeHeaderFooter</param>
		/// <param name="pageDesign">string PageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle LetterStyle</param>
		/// <param name="letterhead">bool Letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation LetterheadLocation</param>
		/// <param name="letterheadSize">Single LetterheadSize</param>
		/// <param name="recipientName">string RecipientName</param>
		/// <param name="recipientAddress">string RecipientAddress</param>
		/// <param name="salutation">string Salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType SalutationType</param>
		/// <param name="recipientReference">string RecipientReference</param>
		/// <param name="mailingInstructions">string MailingInstructions</param>
		/// <param name="attentionLine">string AttentionLine</param>
		/// <param name="subject">string Subject</param>
		/// <param name="cCList">string CCList</param>
		/// <param name="returnAddress">string ReturnAddress</param>
		/// <param name="senderName">string SenderName</param>
		/// <param name="closing">string Closing</param>
		/// <param name="senderCompany">string SenderCompany</param>
		/// <param name="senderJobTitle">string SenderJobTitle</param>
		/// <param name="senderInitials">string SenderInitials</param>
		/// <param name="enclosureNumber">Int32 EnclosureNumber</param>
		/// <param name="infoBlock">optional object InfoBlock</param>
		/// <param name="recipientCode">optional object RecipientCode</param>
		/// <param name="recipientGender">optional object RecipientGender</param>
		/// <param name="returnAddressShortForm">optional object ReturnAddressShortForm</param>
		/// <param name="senderCity">optional object SenderCity</param>
		/// <param name="senderCode">optional object SenderCode</param>
		/// <param name="senderGender">optional object SenderGender</param>
		/// <param name="senderReference">optional object SenderReference</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity, object senderCode, object senderGender, object senderReference)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm, senderCity, senderCode, senderGender, senderReference);
			object returnItem = Invoker.MethodReturn(this, "CreateLetterContent", paramsArray);
			NetOffice.WordApi.LetterContent newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.LetterContent.LateBindingApiWrapperType) as NetOffice.WordApi.LetterContent;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx
		/// </summary>
		/// <param name="dateFormat">string DateFormat</param>
		/// <param name="includeHeaderFooter">bool IncludeHeaderFooter</param>
		/// <param name="pageDesign">string PageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle LetterStyle</param>
		/// <param name="letterhead">bool Letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation LetterheadLocation</param>
		/// <param name="letterheadSize">Single LetterheadSize</param>
		/// <param name="recipientName">string RecipientName</param>
		/// <param name="recipientAddress">string RecipientAddress</param>
		/// <param name="salutation">string Salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType SalutationType</param>
		/// <param name="recipientReference">string RecipientReference</param>
		/// <param name="mailingInstructions">string MailingInstructions</param>
		/// <param name="attentionLine">string AttentionLine</param>
		/// <param name="subject">string Subject</param>
		/// <param name="cCList">string CCList</param>
		/// <param name="returnAddress">string ReturnAddress</param>
		/// <param name="senderName">string SenderName</param>
		/// <param name="closing">string Closing</param>
		/// <param name="senderCompany">string SenderCompany</param>
		/// <param name="senderJobTitle">string SenderJobTitle</param>
		/// <param name="senderInitials">string SenderInitials</param>
		/// <param name="enclosureNumber">Int32 EnclosureNumber</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber);
			object returnItem = Invoker.MethodReturn(this, "CreateLetterContent", paramsArray);
			NetOffice.WordApi.LetterContent newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.LetterContent.LateBindingApiWrapperType) as NetOffice.WordApi.LetterContent;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx
		/// </summary>
		/// <param name="dateFormat">string DateFormat</param>
		/// <param name="includeHeaderFooter">bool IncludeHeaderFooter</param>
		/// <param name="pageDesign">string PageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle LetterStyle</param>
		/// <param name="letterhead">bool Letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation LetterheadLocation</param>
		/// <param name="letterheadSize">Single LetterheadSize</param>
		/// <param name="recipientName">string RecipientName</param>
		/// <param name="recipientAddress">string RecipientAddress</param>
		/// <param name="salutation">string Salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType SalutationType</param>
		/// <param name="recipientReference">string RecipientReference</param>
		/// <param name="mailingInstructions">string MailingInstructions</param>
		/// <param name="attentionLine">string AttentionLine</param>
		/// <param name="subject">string Subject</param>
		/// <param name="cCList">string CCList</param>
		/// <param name="returnAddress">string ReturnAddress</param>
		/// <param name="senderName">string SenderName</param>
		/// <param name="closing">string Closing</param>
		/// <param name="senderCompany">string SenderCompany</param>
		/// <param name="senderJobTitle">string SenderJobTitle</param>
		/// <param name="senderInitials">string SenderInitials</param>
		/// <param name="enclosureNumber">Int32 EnclosureNumber</param>
		/// <param name="infoBlock">optional object InfoBlock</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock);
			object returnItem = Invoker.MethodReturn(this, "CreateLetterContent", paramsArray);
			NetOffice.WordApi.LetterContent newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.LetterContent.LateBindingApiWrapperType) as NetOffice.WordApi.LetterContent;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx
		/// </summary>
		/// <param name="dateFormat">string DateFormat</param>
		/// <param name="includeHeaderFooter">bool IncludeHeaderFooter</param>
		/// <param name="pageDesign">string PageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle LetterStyle</param>
		/// <param name="letterhead">bool Letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation LetterheadLocation</param>
		/// <param name="letterheadSize">Single LetterheadSize</param>
		/// <param name="recipientName">string RecipientName</param>
		/// <param name="recipientAddress">string RecipientAddress</param>
		/// <param name="salutation">string Salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType SalutationType</param>
		/// <param name="recipientReference">string RecipientReference</param>
		/// <param name="mailingInstructions">string MailingInstructions</param>
		/// <param name="attentionLine">string AttentionLine</param>
		/// <param name="subject">string Subject</param>
		/// <param name="cCList">string CCList</param>
		/// <param name="returnAddress">string ReturnAddress</param>
		/// <param name="senderName">string SenderName</param>
		/// <param name="closing">string Closing</param>
		/// <param name="senderCompany">string SenderCompany</param>
		/// <param name="senderJobTitle">string SenderJobTitle</param>
		/// <param name="senderInitials">string SenderInitials</param>
		/// <param name="enclosureNumber">Int32 EnclosureNumber</param>
		/// <param name="infoBlock">optional object InfoBlock</param>
		/// <param name="recipientCode">optional object RecipientCode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode);
			object returnItem = Invoker.MethodReturn(this, "CreateLetterContent", paramsArray);
			NetOffice.WordApi.LetterContent newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.LetterContent.LateBindingApiWrapperType) as NetOffice.WordApi.LetterContent;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx
		/// </summary>
		/// <param name="dateFormat">string DateFormat</param>
		/// <param name="includeHeaderFooter">bool IncludeHeaderFooter</param>
		/// <param name="pageDesign">string PageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle LetterStyle</param>
		/// <param name="letterhead">bool Letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation LetterheadLocation</param>
		/// <param name="letterheadSize">Single LetterheadSize</param>
		/// <param name="recipientName">string RecipientName</param>
		/// <param name="recipientAddress">string RecipientAddress</param>
		/// <param name="salutation">string Salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType SalutationType</param>
		/// <param name="recipientReference">string RecipientReference</param>
		/// <param name="mailingInstructions">string MailingInstructions</param>
		/// <param name="attentionLine">string AttentionLine</param>
		/// <param name="subject">string Subject</param>
		/// <param name="cCList">string CCList</param>
		/// <param name="returnAddress">string ReturnAddress</param>
		/// <param name="senderName">string SenderName</param>
		/// <param name="closing">string Closing</param>
		/// <param name="senderCompany">string SenderCompany</param>
		/// <param name="senderJobTitle">string SenderJobTitle</param>
		/// <param name="senderInitials">string SenderInitials</param>
		/// <param name="enclosureNumber">Int32 EnclosureNumber</param>
		/// <param name="infoBlock">optional object InfoBlock</param>
		/// <param name="recipientCode">optional object RecipientCode</param>
		/// <param name="recipientGender">optional object RecipientGender</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender);
			object returnItem = Invoker.MethodReturn(this, "CreateLetterContent", paramsArray);
			NetOffice.WordApi.LetterContent newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.LetterContent.LateBindingApiWrapperType) as NetOffice.WordApi.LetterContent;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx
		/// </summary>
		/// <param name="dateFormat">string DateFormat</param>
		/// <param name="includeHeaderFooter">bool IncludeHeaderFooter</param>
		/// <param name="pageDesign">string PageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle LetterStyle</param>
		/// <param name="letterhead">bool Letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation LetterheadLocation</param>
		/// <param name="letterheadSize">Single LetterheadSize</param>
		/// <param name="recipientName">string RecipientName</param>
		/// <param name="recipientAddress">string RecipientAddress</param>
		/// <param name="salutation">string Salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType SalutationType</param>
		/// <param name="recipientReference">string RecipientReference</param>
		/// <param name="mailingInstructions">string MailingInstructions</param>
		/// <param name="attentionLine">string AttentionLine</param>
		/// <param name="subject">string Subject</param>
		/// <param name="cCList">string CCList</param>
		/// <param name="returnAddress">string ReturnAddress</param>
		/// <param name="senderName">string SenderName</param>
		/// <param name="closing">string Closing</param>
		/// <param name="senderCompany">string SenderCompany</param>
		/// <param name="senderJobTitle">string SenderJobTitle</param>
		/// <param name="senderInitials">string SenderInitials</param>
		/// <param name="enclosureNumber">Int32 EnclosureNumber</param>
		/// <param name="infoBlock">optional object InfoBlock</param>
		/// <param name="recipientCode">optional object RecipientCode</param>
		/// <param name="recipientGender">optional object RecipientGender</param>
		/// <param name="returnAddressShortForm">optional object ReturnAddressShortForm</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm);
			object returnItem = Invoker.MethodReturn(this, "CreateLetterContent", paramsArray);
			NetOffice.WordApi.LetterContent newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.LetterContent.LateBindingApiWrapperType) as NetOffice.WordApi.LetterContent;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx
		/// </summary>
		/// <param name="dateFormat">string DateFormat</param>
		/// <param name="includeHeaderFooter">bool IncludeHeaderFooter</param>
		/// <param name="pageDesign">string PageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle LetterStyle</param>
		/// <param name="letterhead">bool Letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation LetterheadLocation</param>
		/// <param name="letterheadSize">Single LetterheadSize</param>
		/// <param name="recipientName">string RecipientName</param>
		/// <param name="recipientAddress">string RecipientAddress</param>
		/// <param name="salutation">string Salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType SalutationType</param>
		/// <param name="recipientReference">string RecipientReference</param>
		/// <param name="mailingInstructions">string MailingInstructions</param>
		/// <param name="attentionLine">string AttentionLine</param>
		/// <param name="subject">string Subject</param>
		/// <param name="cCList">string CCList</param>
		/// <param name="returnAddress">string ReturnAddress</param>
		/// <param name="senderName">string SenderName</param>
		/// <param name="closing">string Closing</param>
		/// <param name="senderCompany">string SenderCompany</param>
		/// <param name="senderJobTitle">string SenderJobTitle</param>
		/// <param name="senderInitials">string SenderInitials</param>
		/// <param name="enclosureNumber">Int32 EnclosureNumber</param>
		/// <param name="infoBlock">optional object InfoBlock</param>
		/// <param name="recipientCode">optional object RecipientCode</param>
		/// <param name="recipientGender">optional object RecipientGender</param>
		/// <param name="returnAddressShortForm">optional object ReturnAddressShortForm</param>
		/// <param name="senderCity">optional object SenderCity</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm, senderCity);
			object returnItem = Invoker.MethodReturn(this, "CreateLetterContent", paramsArray);
			NetOffice.WordApi.LetterContent newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.LetterContent.LateBindingApiWrapperType) as NetOffice.WordApi.LetterContent;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx
		/// </summary>
		/// <param name="dateFormat">string DateFormat</param>
		/// <param name="includeHeaderFooter">bool IncludeHeaderFooter</param>
		/// <param name="pageDesign">string PageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle LetterStyle</param>
		/// <param name="letterhead">bool Letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation LetterheadLocation</param>
		/// <param name="letterheadSize">Single LetterheadSize</param>
		/// <param name="recipientName">string RecipientName</param>
		/// <param name="recipientAddress">string RecipientAddress</param>
		/// <param name="salutation">string Salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType SalutationType</param>
		/// <param name="recipientReference">string RecipientReference</param>
		/// <param name="mailingInstructions">string MailingInstructions</param>
		/// <param name="attentionLine">string AttentionLine</param>
		/// <param name="subject">string Subject</param>
		/// <param name="cCList">string CCList</param>
		/// <param name="returnAddress">string ReturnAddress</param>
		/// <param name="senderName">string SenderName</param>
		/// <param name="closing">string Closing</param>
		/// <param name="senderCompany">string SenderCompany</param>
		/// <param name="senderJobTitle">string SenderJobTitle</param>
		/// <param name="senderInitials">string SenderInitials</param>
		/// <param name="enclosureNumber">Int32 EnclosureNumber</param>
		/// <param name="infoBlock">optional object InfoBlock</param>
		/// <param name="recipientCode">optional object RecipientCode</param>
		/// <param name="recipientGender">optional object RecipientGender</param>
		/// <param name="returnAddressShortForm">optional object ReturnAddressShortForm</param>
		/// <param name="senderCity">optional object SenderCity</param>
		/// <param name="senderCode">optional object SenderCode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity, object senderCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm, senderCity, senderCode);
			object returnItem = Invoker.MethodReturn(this, "CreateLetterContent", paramsArray);
			NetOffice.WordApi.LetterContent newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.LetterContent.LateBindingApiWrapperType) as NetOffice.WordApi.LetterContent;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx
		/// </summary>
		/// <param name="dateFormat">string DateFormat</param>
		/// <param name="includeHeaderFooter">bool IncludeHeaderFooter</param>
		/// <param name="pageDesign">string PageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle LetterStyle</param>
		/// <param name="letterhead">bool Letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation LetterheadLocation</param>
		/// <param name="letterheadSize">Single LetterheadSize</param>
		/// <param name="recipientName">string RecipientName</param>
		/// <param name="recipientAddress">string RecipientAddress</param>
		/// <param name="salutation">string Salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType SalutationType</param>
		/// <param name="recipientReference">string RecipientReference</param>
		/// <param name="mailingInstructions">string MailingInstructions</param>
		/// <param name="attentionLine">string AttentionLine</param>
		/// <param name="subject">string Subject</param>
		/// <param name="cCList">string CCList</param>
		/// <param name="returnAddress">string ReturnAddress</param>
		/// <param name="senderName">string SenderName</param>
		/// <param name="closing">string Closing</param>
		/// <param name="senderCompany">string SenderCompany</param>
		/// <param name="senderJobTitle">string SenderJobTitle</param>
		/// <param name="senderInitials">string SenderInitials</param>
		/// <param name="enclosureNumber">Int32 EnclosureNumber</param>
		/// <param name="infoBlock">optional object InfoBlock</param>
		/// <param name="recipientCode">optional object RecipientCode</param>
		/// <param name="recipientGender">optional object RecipientGender</param>
		/// <param name="returnAddressShortForm">optional object ReturnAddressShortForm</param>
		/// <param name="senderCity">optional object SenderCity</param>
		/// <param name="senderCode">optional object SenderCode</param>
		/// <param name="senderGender">optional object SenderGender</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity, object senderCode, object senderGender)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm, senderCity, senderCode, senderGender);
			object returnItem = Invoker.MethodReturn(this, "CreateLetterContent", paramsArray);
			NetOffice.WordApi.LetterContent newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.LetterContent.LateBindingApiWrapperType) as NetOffice.WordApi.LetterContent;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193342.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AcceptAllRevisions()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AcceptAllRevisions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838536.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RejectAllRevisions()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RejectAllRevisions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197127.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DetectLanguage()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DetectLanguage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835740.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ApplyTheme(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "ApplyTheme", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839088.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RemoveTheme()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RemoveTheme", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835177.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void WebPagePreview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "WebPagePreview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195768.aspx
		/// </summary>
		/// <param name="encoding">NetOffice.OfficeApi.Enums.MsoEncoding Encoding</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(encoding);
			Invoker.Method(this, "ReloadAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object PrintZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object PrintZoomPaperHeight</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object PrintZoomPaperWidth</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="s">string s</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void sblt(string s)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(s);
			Invoker.Method(this, "sblt", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter);
			Invoker.Method(this, "SaveAs2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs2000()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SaveAs2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "SaveAs2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat);
			Invoker.Method(this, "SaveAs2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments);
			Invoker.Method(this, "SaveAs2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password);
			Invoker.Method(this, "SaveAs2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles);
			Invoker.Method(this, "SaveAs2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword);
			Invoker.Method(this, "SaveAs2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended);
			Invoker.Method(this, "SaveAs2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts);
			Invoker.Method(this, "SaveAs2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat);
			Invoker.Method(this, "SaveAs2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData);
			Invoker.Method(this, "SaveAs2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Compare2000(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "Compare2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Merge2000(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "Merge2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object PrintZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object PrintZoomPaperHeight</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object PrintZoomPaperWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838511.aspx
		/// </summary>
		/// <param name="codePageOrigin">Int32 CodePageOrigin</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ConvertVietDoc(Int32 codePageOrigin)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(codePageOrigin);
			Invoker.Method(this, "ConvertVietDoc", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194029.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments, object makePublic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194029.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void CheckIn()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194029.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194029.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198206.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool CanCheckin()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CanCheckin", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		/// <param name="subject">optional object Subject</param>
		/// <param name="showMessage">optional object ShowMessage</param>
		/// <param name="includeAttachment">optional object IncludeAttachment</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject, object showMessage, object includeAttachment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, showMessage, includeAttachment);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SendForReview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SendForReview(object recipients)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		/// <param name="subject">optional object Subject</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		/// <param name="subject">optional object Subject</param>
		/// <param name="showMessage">optional object ShowMessage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject, object showMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, showMessage);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836324.aspx
		/// </summary>
		/// <param name="showMessage">optional object ShowMessage</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ReplyWithChanges(object showMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showMessage);
			Invoker.Method(this, "ReplyWithChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836324.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ReplyWithChanges()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ReplyWithChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837660.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void EndReview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EndReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195460.aspx
		/// </summary>
		/// <param name="passwordEncryptionProvider">string PasswordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string PasswordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 PasswordEncryptionKeyLength</param>
		/// <param name="passwordEncryptionFileProperties">optional object PasswordEncryptionFileProperties</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength, object passwordEncryptionFileProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength, passwordEncryptionFileProperties);
			Invoker.Method(this, "SetPasswordEncryptionOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195460.aspx
		/// </summary>
		/// <param name="passwordEncryptionProvider">string PasswordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string PasswordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 PasswordEncryptionKeyLength</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength);
			Invoker.Method(this, "SetPasswordEncryptionOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void RecheckSmartTags()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RecheckSmartTags", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void RemoveSmartTags()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RemoveSmartTags", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198118.aspx
		/// </summary>
		/// <param name="style">object Style</param>
		/// <param name="setInTemplate">bool SetInTemplate</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SetDefaultTableStyle(object style, bool setInTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, setInTemplate);
			Invoker.Method(this, "SetDefaultTableStyle", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822910.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void DeleteAllComments()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DeleteAllComments", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837501.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void AcceptAllRevisionsShown()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AcceptAllRevisionsShown", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822533.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void RejectAllRevisionsShown()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RejectAllRevisionsShown", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836620.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void DeleteAllCommentsShown()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DeleteAllCommentsShown", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821137.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ResetFormFields()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ResetFormFields", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void CheckNewSmartTags()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckNewSmartTags", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType Type</param>
		/// <param name="noReset">optional object NoReset</param>
		/// <param name="password">optional object Password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Protect2002(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, noReset, password);
			Invoker.Method(this, "Protect2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType Type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Protect2002(NetOffice.WordApi.Enums.WdProtectionType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "Protect2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType Type</param>
		/// <param name="noReset">optional object NoReset</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Protect2002(NetOffice.WordApi.Enums.WdProtectionType type, object noReset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, noReset);
			Invoker.Method(this, "Protect2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="authorName">optional object AuthorName</param>
		/// <param name="compareTarget">optional object CompareTarget</param>
		/// <param name="detectFormatChanges">optional object DetectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object IgnoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Compare2002(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings, addToRecentFiles);
			Invoker.Method(this, "Compare2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Compare2002(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "Compare2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="authorName">optional object AuthorName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Compare2002(string name, object authorName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, authorName);
			Invoker.Method(this, "Compare2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="authorName">optional object AuthorName</param>
		/// <param name="compareTarget">optional object CompareTarget</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Compare2002(string name, object authorName, object compareTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, authorName, compareTarget);
			Invoker.Method(this, "Compare2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="authorName">optional object AuthorName</param>
		/// <param name="compareTarget">optional object CompareTarget</param>
		/// <param name="detectFormatChanges">optional object DetectFormatChanges</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Compare2002(string name, object authorName, object compareTarget, object detectFormatChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, authorName, compareTarget, detectFormatChanges);
			Invoker.Method(this, "Compare2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="authorName">optional object AuthorName</param>
		/// <param name="compareTarget">optional object CompareTarget</param>
		/// <param name="detectFormatChanges">optional object DetectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object IgnoreAllComparisonWarnings</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Compare2002(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings);
			Invoker.Method(this, "Compare2002", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192018.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		/// <param name="subject">optional object Subject</param>
		/// <param name="showMessage">optional object ShowMessage</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject, object showMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, showMessage);
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192018.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void SendFaxOverInternet()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192018.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients);
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192018.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		/// <param name="subject">optional object Subject</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject);
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196274.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="dataOnly">optional bool DataOnly = true</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void TransformDocument(string path, object dataOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, dataOnly);
			Invoker.Method(this, "TransformDocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196274.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void TransformDocument(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(this, "TransformDocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195660.aspx
		/// </summary>
		/// <param name="editorID">optional object EditorID</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void SelectAllEditableRanges(object editorID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(editorID);
			Invoker.Method(this, "SelectAllEditableRanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195660.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void SelectAllEditableRanges()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectAllEditableRanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844883.aspx
		/// </summary>
		/// <param name="editorID">optional object EditorID</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void DeleteAllEditableRanges(object editorID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(editorID);
			Invoker.Method(this, "DeleteAllEditableRanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844883.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void DeleteAllEditableRanges()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DeleteAllEditableRanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838947.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void DeleteAllInkAnnotations()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DeleteAllInkAnnotations", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="richFormat">bool RichFormat</param>
		/// <param name="url">string Url</param>
		/// <param name="title">string Title</param>
		/// <param name="description">string Description</param>
		/// <param name="iD">string ID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void AddDocumentWorkspaceHeader(bool richFormat, string url, string title, string description, string iD)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(richFormat, url, title, description, iD);
			Invoker.Method(this, "AddDocumentWorkspaceHeader", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="iD">string ID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void RemoveDocumentWorkspaceHeader(string iD)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iD);
			Invoker.Method(this, "RemoveDocumentWorkspaceHeader", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845389.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void RemoveLockedStyles()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RemoveLockedStyles", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822346.aspx
		/// </summary>
		/// <param name="xPath">string XPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		/// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping, object fastSearchSkippingTextNodes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath, prefixMapping, fastSearchSkippingTextNodes);
			object returnItem = Invoker.MethodReturn(this, "SelectSingleNode", paramsArray);
			NetOffice.WordApi.XMLNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.XMLNode.LateBindingApiWrapperType) as NetOffice.WordApi.XMLNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822346.aspx
		/// </summary>
		/// <param name="xPath">string XPath</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode SelectSingleNode(string xPath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath);
			object returnItem = Invoker.MethodReturn(this, "SelectSingleNode", paramsArray);
			NetOffice.WordApi.XMLNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.XMLNode.LateBindingApiWrapperType) as NetOffice.WordApi.XMLNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822346.aspx
		/// </summary>
		/// <param name="xPath">string XPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath, prefixMapping);
			object returnItem = Invoker.MethodReturn(this, "SelectSingleNode", paramsArray);
			NetOffice.WordApi.XMLNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.XMLNode.LateBindingApiWrapperType) as NetOffice.WordApi.XMLNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837252.aspx
		/// </summary>
		/// <param name="xPath">string XPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		/// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping, object fastSearchSkippingTextNodes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath, prefixMapping, fastSearchSkippingTextNodes);
			object returnItem = Invoker.MethodReturn(this, "SelectNodes", paramsArray);
			NetOffice.WordApi.XMLNodes newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType) as NetOffice.WordApi.XMLNodes;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837252.aspx
		/// </summary>
		/// <param name="xPath">string XPath</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes SelectNodes(string xPath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath);
			object returnItem = Invoker.MethodReturn(this, "SelectNodes", paramsArray);
			NetOffice.WordApi.XMLNodes newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType) as NetOffice.WordApi.XMLNodes;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837252.aspx
		/// </summary>
		/// <param name="xPath">string XPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath, prefixMapping);
			object returnItem = Invoker.MethodReturn(this, "SelectNodes", paramsArray);
			NetOffice.WordApi.XMLNodes newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType) as NetOffice.WordApi.XMLNodes;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197270.aspx
		/// </summary>
		/// <param name="removeDocInfoType">NetOffice.WordApi.Enums.WdRemoveDocInfoType RemoveDocInfoType</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void RemoveDocumentInformation(NetOffice.WordApi.Enums.WdRemoveDocInfoType removeDocInfoType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(removeDocInfoType);
			Invoker.Method(this, "RemoveDocumentInformation", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		/// <param name="versionType">optional object VersionType</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic, versionType);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void CheckInWithVersion()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void Dummy2()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845518.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void LockServerFile()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "LockServerFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198071.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetWorkflowTasks", paramsArray);
			NetOffice.OfficeApi.WorkflowTasks newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.WorkflowTasks.LateBindingApiWrapperType) as NetOffice.OfficeApi.WorkflowTasks;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845242.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetWorkflowTemplates", paramsArray);
			NetOffice.OfficeApi.WorkflowTemplates newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.WorkflowTemplates.LateBindingApiWrapperType) as NetOffice.OfficeApi.WorkflowTemplates;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void Dummy4()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy4", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="skipIfAbsent">bool SkipIfAbsent</param>
		/// <param name="url">string Url</param>
		/// <param name="title">string Title</param>
		/// <param name="description">string Description</param>
		/// <param name="iD">string ID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void AddMeetingWorkspaceHeader(bool skipIfAbsent, string url, string title, string description, string iD)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(skipIfAbsent, url, title, description, iD);
			Invoker.Method(this, "AddMeetingWorkspaceHeader", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198291.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void SaveAsQuickStyleSet(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "SaveAsQuickStyleSet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ApplyQuickStyleSet(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "ApplyQuickStyleSet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840910.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ApplyDocumentTheme(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "ApplyDocumentTheme", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838276.aspx
		/// </summary>
		/// <param name="node">NetOffice.OfficeApi.CustomXMLNode Node</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControls SelectLinkedControls(NetOffice.OfficeApi.CustomXMLNode node)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(node);
			object returnItem = Invoker.MethodReturn(this, "SelectLinkedControls", paramsArray);
			NetOffice.WordApi.ContentControls newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.ContentControls.LateBindingApiWrapperType) as NetOffice.WordApi.ContentControls;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198010.aspx
		/// </summary>
		/// <param name="stream">optional NetOffice.OfficeApi.CustomXMLPart Stream = 0</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControls SelectUnlinkedControls(object stream)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(stream);
			object returnItem = Invoker.MethodReturn(this, "SelectUnlinkedControls", paramsArray);
			NetOffice.WordApi.ContentControls newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.ContentControls.LateBindingApiWrapperType) as NetOffice.WordApi.ContentControls;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198010.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControls SelectUnlinkedControls()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "SelectUnlinkedControls", paramsArray);
			NetOffice.WordApi.ContentControls newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.ContentControls.LateBindingApiWrapperType) as NetOffice.WordApi.ContentControls;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822990.aspx
		/// </summary>
		/// <param name="title">string Title</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControls SelectContentControlsByTitle(string title)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(title);
			object returnItem = Invoker.MethodReturn(this, "SelectContentControlsByTitle", paramsArray);
			NetOffice.WordApi.ContentControls newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.ContentControls.LateBindingApiWrapperType) as NetOffice.WordApi.ContentControls;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
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
		/// <param name="fixedFormatExtClassPtr">optional object FixedFormatExtClassPtr</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object fixedFormatExtClassPtr)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1, fixedFormatExtClassPtr);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, range);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, range, from);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
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
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
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
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx
		/// </summary>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat ExportFormat</param>
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
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196504.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void FreezeLayout()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FreezeLayout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void UnfreezeLayout()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UnfreezeLayout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194276.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void DowngradeDocument()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DowngradeDocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835714.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void Convert()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Convert", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839693.aspx
		/// </summary>
		/// <param name="tag">string Tag</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControls SelectContentControlsByTag(string tag)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tag);
			object returnItem = Invoker.MethodReturn(this, "SelectContentControlsByTag", paramsArray);
			NetOffice.WordApi.ContentControls newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.ContentControls.LateBindingApiWrapperType) as NetOffice.WordApi.ContentControls;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838360.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ConvertAutoHyphens()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ConvertAutoHyphens", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821672.aspx
		/// </summary>
		/// <param name="style">object Style</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyQuickStyleSet2(object style)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style);
			Invoker.Method(this, "ApplyQuickStyleSet2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object AllowSubstitutions</param>
		/// <param name="lineEnding">optional object LineEnding</param>
		/// <param name="addBiDiMarks">optional object AddBiDiMarks</param>
		/// <param name="compatibilityMode">optional object CompatibilityMode</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks, object compatibilityMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks, compatibilityMode);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object AllowSubstitutions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object AllowSubstitutions</param>
		/// <param name="lineEnding">optional object LineEnding</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object AllowSubstitutions</param>
		/// <param name="lineEnding">optional object LineEnding</param>
		/// <param name="addBiDiMarks">optional object AddBiDiMarks</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks);
			Invoker.Method(this, "SaveAs2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840359.aspx
		/// </summary>
		/// <param name="mode">Int32 Mode</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SetCompatibilityMode(Int32 mode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(mode);
			Invoker.Method(this, "SetCompatibilityMode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231927.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 15, 16)]
		public Int32 ReturnToLastReadPosition()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ReturnToLastReadPosition", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object AllowSubstitutions</param>
		/// <param name="lineEnding">optional object LineEnding</param>
		/// <param name="addBiDiMarks">optional object AddBiDiMarks</param>
		/// <param name="compatibilityMode">optional object CompatibilityMode</param>
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks, object compatibilityMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks, compatibilityMode);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object AllowSubstitutions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object AllowSubstitutions</param>
		/// <param name="lineEnding">optional object LineEnding</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="lockComments">optional object LockComments</param>
		/// <param name="password">optional object Password</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object EmbedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object SaveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object SaveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object SaveAsAOCELetter</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="insertLineBreaks">optional object InsertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object AllowSubstitutions</param>
		/// <param name="lineEnding">optional object LineEnding</param>
		/// <param name="addBiDiMarks">optional object AddBiDiMarks</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}