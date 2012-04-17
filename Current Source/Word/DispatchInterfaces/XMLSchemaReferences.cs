using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using LateBindingApi.Core;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface XMLSchemaReferences 
	/// SupportByVersion Word, 11,12,14
	///</summary>
	[SupportByVersionAttribute("Word", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class XMLSchemaReferences : COMObject ,IEnumerable<NetOffice.WordApi.XMLSchemaReference>
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
                    _type = typeof(XMLSchemaReferences);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLSchemaReferences(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLSchemaReferences(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLSchemaReferences(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLSchemaReferences() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLSchemaReferences(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
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
		/// SupportByVersion Word 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
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
		/// SupportByVersion Word 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
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
		/// SupportByVersion Word 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
		public bool AutomaticValidation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutomaticValidation", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutomaticValidation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
		public bool AllowSaveAsXMLWithoutValidation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AllowSaveAsXMLWithoutValidation", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AllowSaveAsXMLWithoutValidation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
		public bool HideValidationErrors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HideValidationErrors", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HideValidationErrors", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
		public bool IgnoreMixedContent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IgnoreMixedContent", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IgnoreMixedContent", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
		public bool ShowPlaceholderText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowPlaceholderText", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowPlaceholderText", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Word", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.WordApi.XMLSchemaReference this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.WordApi.XMLSchemaReference newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.XMLSchemaReference.LateBindingApiWrapperType) as NetOffice.WordApi.XMLSchemaReference;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
		public void Validate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Validate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// </summary>
		/// <param name="namespaceURI">object NamespaceURI</param>
		/// <param name="alias">object Alias</param>
		/// <param name="fileName">object FileName</param>
		/// <param name="installForAllUsers">optional bool InstallForAllUsers = false</param>
		[SupportByVersionAttribute("Word", 11,12,14)]
		public NetOffice.WordApi.XMLSchemaReference Add(object namespaceURI, object alias, object fileName, bool installForAllUsers)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(namespaceURI, alias, fileName, installForAllUsers);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.XMLSchemaReference newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.XMLSchemaReference.LateBindingApiWrapperType) as NetOffice.WordApi.XMLSchemaReference;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// </summary>
		/// <param name="namespaceURI">object NamespaceURI</param>
		/// <param name="alias">object Alias</param>
		/// <param name="fileName">object FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14)]
		public NetOffice.WordApi.XMLSchemaReference Add(object namespaceURI, object alias, object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(namespaceURI, alias, fileName);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.XMLSchemaReference newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.XMLSchemaReference.LateBindingApiWrapperType) as NetOffice.WordApi.XMLSchemaReference;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.WordApi.XMLSchemaReference> Member
        
        /// <summary>
		/// SupportByVersionAttribute Word, 11,12,14
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
       public IEnumerator<NetOffice.WordApi.XMLSchemaReference> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.WordApi.XMLSchemaReference item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Word, 11,12,14
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return LateBindingApi.Core.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}