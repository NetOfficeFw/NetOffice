using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface _CustomXMLSchemaCollection 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class _CustomXMLSchemaCollection : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.CustomXMLSchema>
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
                    _type = typeof(_CustomXMLSchemaCollection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _CustomXMLSchemaCollection(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _CustomXMLSchemaCollection(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLSchemaCollection(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLSchemaCollection(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLSchemaCollection(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLSchemaCollection(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLSchemaCollection() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLSchemaCollection(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLSchemaCollection.Parent"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLSchemaCollection.Count"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Office", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.CustomXMLSchema this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLSchema>(this, "Item", NetOffice.OfficeApi.CustomXMLSchema.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLSchemaCollection.NamespaceURI"/> </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_NamespaceURI(Int32 index)
		{
			return Factory.ExecuteStringPropertyGet(this, "NamespaceURI", index);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_NamespaceURI
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLSchemaCollection.NamespaceURI"/> </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_NamespaceURI")]
		public string NamespaceURI(Int32 index)
		{
			return get_NamespaceURI(index);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLSchemaCollection.Add"/> </remarks>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="alias">optional string Alias = </param>
		/// <param name="fileName">optional string FileName = </param>
		/// <param name="installForAllUsers">optional bool InstallForAllUsers = false</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLSchema Add(object namespaceURI, object alias, object fileName, object installForAllUsers)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLSchema>(this, "Add", NetOffice.OfficeApi.CustomXMLSchema.LateBindingApiWrapperType, namespaceURI, alias, fileName, installForAllUsers);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLSchemaCollection.Add"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLSchema Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLSchema>(this, "Add", NetOffice.OfficeApi.CustomXMLSchema.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLSchemaCollection.Add"/> </remarks>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLSchema Add(object namespaceURI)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLSchema>(this, "Add", NetOffice.OfficeApi.CustomXMLSchema.LateBindingApiWrapperType, namespaceURI);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLSchemaCollection.Add"/> </remarks>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="alias">optional string Alias = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLSchema Add(object namespaceURI, object alias)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLSchema>(this, "Add", NetOffice.OfficeApi.CustomXMLSchema.LateBindingApiWrapperType, namespaceURI, alias);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLSchemaCollection.Add"/> </remarks>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="alias">optional string Alias = </param>
		/// <param name="fileName">optional string FileName = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLSchema Add(object namespaceURI, object alias, object fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLSchema>(this, "Add", NetOffice.OfficeApi.CustomXMLSchema.LateBindingApiWrapperType, namespaceURI, alias, fileName);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLSchemaCollection.AddCollection"/> </remarks>
		/// <param name="schemaCollection">NetOffice.OfficeApi.CustomXMLSchemaCollection schemaCollection</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void AddCollection(NetOffice.OfficeApi.CustomXMLSchemaCollection schemaCollection)
		{
			 Factory.ExecuteMethod(this, "AddCollection", schemaCollection);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLSchemaCollection.Validate"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool Validate()
		{
			return Factory.ExecuteBoolMethodGet(this, "Validate");
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.CustomXMLSchema>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.CustomXMLSchema>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.CustomXMLSchema>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.CustomXMLSchema>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.CustomXMLSchema> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.CustomXMLSchema item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}