using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface XMLSchemaReferences 
	/// SupportByVersion Word, 11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences"/> </remarks>
	[SupportByVersion("Word", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class XMLSchemaReferences : COMObject, IEnumerableProvider<NetOffice.WordApi.XMLSchemaReference>
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
                    _type = typeof(XMLSchemaReferences);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public XMLSchemaReferences(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public XMLSchemaReferences(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLSchemaReferences(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLSchemaReferences(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLSchemaReferences(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLSchemaReferences(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLSchemaReferences() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLSchemaReferences(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.Count"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.Application"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.Creator"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.Parent"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool AutomaticValidation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutomaticValidation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutomaticValidation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool AllowSaveAsXMLWithoutValidation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowSaveAsXMLWithoutValidation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowSaveAsXMLWithoutValidation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.HideValidationErrors"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool HideValidationErrors
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HideValidationErrors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HideValidationErrors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.IgnoreMixedContent"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool IgnoreMixedContent
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IgnoreMixedContent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IgnoreMixedContent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.ShowPlaceholderText"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool ShowPlaceholderText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowPlaceholderText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowPlaceholderText", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.WordApi.XMLSchemaReference this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLSchemaReference>(this, "Item", NetOffice.WordApi.XMLSchemaReference.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.Validate"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Validate()
		{
			 Factory.ExecuteMethod(this, "Validate");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.Add"/> </remarks>
		/// <param name="namespaceURI">optional object namespaceURI</param>
		/// <param name="alias">optional object alias</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="installForAllUsers">optional bool InstallForAllUsers = false</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLSchemaReference Add(object namespaceURI, object alias, object fileName, object installForAllUsers)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLSchemaReference>(this, "Add", NetOffice.WordApi.XMLSchemaReference.LateBindingApiWrapperType, namespaceURI, alias, fileName, installForAllUsers);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.Add"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLSchemaReference Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLSchemaReference>(this, "Add", NetOffice.WordApi.XMLSchemaReference.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.Add"/> </remarks>
		/// <param name="namespaceURI">optional object namespaceURI</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLSchemaReference Add(object namespaceURI)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLSchemaReference>(this, "Add", NetOffice.WordApi.XMLSchemaReference.LateBindingApiWrapperType, namespaceURI);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.Add"/> </remarks>
		/// <param name="namespaceURI">optional object namespaceURI</param>
		/// <param name="alias">optional object alias</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLSchemaReference Add(object namespaceURI, object alias)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLSchemaReference>(this, "Add", NetOffice.WordApi.XMLSchemaReference.LateBindingApiWrapperType, namespaceURI, alias);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLSchemaReferences.Add"/> </remarks>
		/// <param name="namespaceURI">optional object namespaceURI</param>
		/// <param name="alias">optional object alias</param>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLSchemaReference Add(object namespaceURI, object alias, object fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLSchemaReference>(this, "Add", NetOffice.WordApi.XMLSchemaReference.LateBindingApiWrapperType, namespaceURI, alias, fileName);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.XMLSchemaReference>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.XMLSchemaReference>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.XMLSchemaReference>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.XMLSchemaReference>

        /// <summary>
        /// SupportByVersion Word, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.WordApi.XMLSchemaReference> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.XMLSchemaReference item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Word, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}