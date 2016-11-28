using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface Documents 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840891.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Documents : COMObject ,IEnumerable<NetOffice.WordApi.Document>
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
                    _type = typeof(Documents);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Documents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Documents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Documents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Documents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Documents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Documents() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Documents(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822958.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195113.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838145.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196684.aspx
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.WordApi.Document this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844896.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844896.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844896.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844896.aspx
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
		/// <param name="template">optional object Template</param>
		/// <param name="newTemplate">optional object NewTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document AddOld(object template, object newTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(template, newTemplate);
			object returnItem = Invoker.MethodReturn(this, "AddOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document AddOld()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="template">optional object Template</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document AddOld(object template)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(template);
			object returnItem = Invoker.MethodReturn(this, "AddOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format);
			object returnItem = Invoker.MethodReturn(this, "OpenOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "OpenOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions);
			object returnItem = Invoker.MethodReturn(this, "OpenOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly);
			object returnItem = Invoker.MethodReturn(this, "OpenOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles);
			object returnItem = Invoker.MethodReturn(this, "OpenOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument);
			object returnItem = Invoker.MethodReturn(this, "OpenOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate);
			object returnItem = Invoker.MethodReturn(this, "OpenOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert);
			object returnItem = Invoker.MethodReturn(this, "OpenOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument);
			object returnItem = Invoker.MethodReturn(this, "OpenOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate);
			object returnItem = Invoker.MethodReturn(this, "OpenOld", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195961.aspx
		/// </summary>
		/// <param name="noPrompt">optional object NoPrompt</param>
		/// <param name="originalFormat">optional object OriginalFormat</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Save(object noPrompt, object originalFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(noPrompt, originalFormat);
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195961.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Save()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195961.aspx
		/// </summary>
		/// <param name="noPrompt">optional object NoPrompt</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Save(object noPrompt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(noPrompt);
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845011.aspx
		/// </summary>
		/// <param name="template">optional object Template</param>
		/// <param name="newTemplate">optional object NewTemplate</param>
		/// <param name="documentType">optional object DocumentType</param>
		/// <param name="visible">optional object Visible</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Add(object template, object newTemplate, object documentType, object visible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(template, newTemplate, documentType, visible);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845011.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Add()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845011.aspx
		/// </summary>
		/// <param name="template">optional object Template</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Add(object template)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(template);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845011.aspx
		/// </summary>
		/// <param name="template">optional object Template</param>
		/// <param name="newTemplate">optional object NewTemplate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Add(object template, object newTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(template, newTemplate);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845011.aspx
		/// </summary>
		/// <param name="template">optional object Template</param>
		/// <param name="newTemplate">optional object NewTemplate</param>
		/// <param name="documentType">optional object DocumentType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Add(object template, object newTemplate, object documentType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(template, newTemplate, documentType);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="openAndRepair">optional object OpenAndRepair</param>
		/// <param name="documentDirection">optional object DocumentDirection</param>
		/// <param name="noEncodingDialog">optional object NoEncodingDialog</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection, noEncodingDialog);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="openAndRepair">optional object OpenAndRepair</param>
		/// <param name="documentDirection">optional object DocumentDirection</param>
		/// <param name="noEncodingDialog">optional object NoEncodingDialog</param>
		/// <param name="xMLTransform">optional object XMLTransform</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog, object xMLTransform)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection, noEncodingDialog, xMLTransform);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="openAndRepair">optional object OpenAndRepair</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835182.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="openAndRepair">optional object OpenAndRepair</param>
		/// <param name="documentDirection">optional object DocumentDirection</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible);
			object returnItem = Invoker.MethodReturn(this, "Open2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "Open2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions);
			object returnItem = Invoker.MethodReturn(this, "Open2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly);
			object returnItem = Invoker.MethodReturn(this, "Open2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles);
			object returnItem = Invoker.MethodReturn(this, "Open2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument);
			object returnItem = Invoker.MethodReturn(this, "Open2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate);
			object returnItem = Invoker.MethodReturn(this, "Open2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert);
			object returnItem = Invoker.MethodReturn(this, "Open2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument);
			object returnItem = Invoker.MethodReturn(this, "Open2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate);
			object returnItem = Invoker.MethodReturn(this, "Open2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format);
			object returnItem = Invoker.MethodReturn(this, "Open2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding);
			object returnItem = Invoker.MethodReturn(this, "Open2000", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198275.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void CheckOut(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "CheckOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839907.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool CanCheckOut(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "CanCheckOut", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="openAndRepair">optional object OpenAndRepair</param>
		/// <param name="documentDirection">optional object DocumentDirection</param>
		/// <param name="noEncodingDialog">optional object NoEncodingDialog</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection, noEncodingDialog);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="openAndRepair">optional object OpenAndRepair</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="openAndRepair">optional object OpenAndRepair</param>
		/// <param name="documentDirection">optional object DocumentDirection</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection);
			object returnItem = Invoker.MethodReturn(this, "Open2002", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="openAndRepair">optional object OpenAndRepair</param>
		/// <param name="documentDirection">optional object DocumentDirection</param>
		/// <param name="noEncodingDialog">optional object NoEncodingDialog</param>
		/// <param name="xMLTransform">optional object XMLTransform</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog, object xMLTransform)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection, noEncodingDialog, xMLTransform);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="openAndRepair">optional object OpenAndRepair</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="openAndRepair">optional object OpenAndRepair</param>
		/// <param name="documentDirection">optional object DocumentDirection</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839499.aspx
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="format">optional object Format</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="openAndRepair">optional object OpenAndRepair</param>
		/// <param name="documentDirection">optional object DocumentDirection</param>
		/// <param name="noEncodingDialog">optional object NoEncodingDialog</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection, noEncodingDialog);
			object returnItem = Invoker.MethodReturn(this, "OpenNoRepairDialog", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838738.aspx
		/// </summary>
		/// <param name="providerID">string ProviderID</param>
		/// <param name="postURL">string PostURL</param>
		/// <param name="blogName">string BlogName</param>
		/// <param name="postID">optional string PostID = </param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document AddBlogDocument(string providerID, string postURL, string blogName, object postID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(providerID, postURL, blogName, postID);
			object returnItem = Invoker.MethodReturn(this, "AddBlogDocument", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838738.aspx
		/// </summary>
		/// <param name="providerID">string ProviderID</param>
		/// <param name="postURL">string PostURL</param>
		/// <param name="blogName">string BlogName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document AddBlogDocument(string providerID, string postURL, string blogName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(providerID, postURL, blogName);
			object returnItem = Invoker.MethodReturn(this, "AddBlogDocument", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.WordApi.Document> Member
        
        /// <summary>
		/// SupportByVersionAttribute Word, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
       public IEnumerator<NetOffice.WordApi.Document> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.WordApi.Document item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Word, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}