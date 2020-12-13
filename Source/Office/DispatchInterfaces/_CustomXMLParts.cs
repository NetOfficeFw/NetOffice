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
	/// DispatchInterface _CustomXMLParts 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class _CustomXMLParts : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.CustomXMLPart>
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
                    _type = typeof(_CustomXMLParts);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _CustomXMLParts(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _CustomXMLParts(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLParts(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLParts(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLParts(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLParts(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLParts() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLParts(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLParts.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLParts.Count"/> </remarks>
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
		public NetOffice.OfficeApi.CustomXMLPart this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLPart>(this, "Item", NetOffice.OfficeApi.CustomXMLPart.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLParts.Add"/> </remarks>
		/// <param name="xML">optional string XML = </param>
		/// <param name="schemaCollection">optional object schemaCollection</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLPart Add(object xML, object schemaCollection)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLPart>(this, "Add", NetOffice.OfficeApi.CustomXMLPart.LateBindingApiWrapperType, xML, schemaCollection);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLParts.Add"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLPart Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLPart>(this, "Add", NetOffice.OfficeApi.CustomXMLPart.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLParts.Add"/> </remarks>
		/// <param name="xML">optional string XML = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLPart Add(object xML)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLPart>(this, "Add", NetOffice.OfficeApi.CustomXMLPart.LateBindingApiWrapperType, xML);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLParts.SelectByID"/> </remarks>
		/// <param name="id">string id</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLPart SelectByID(string id)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLPart>(this, "SelectByID", NetOffice.OfficeApi.CustomXMLPart.LateBindingApiWrapperType, id);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLParts.SelectByNamespace"/> </remarks>
		/// <param name="namespaceURI">string namespaceURI</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLParts SelectByNamespace(string namespaceURI)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLParts>(this, "SelectByNamespace", NetOffice.OfficeApi.CustomXMLParts.LateBindingApiWrapperType, namespaceURI);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.CustomXMLPart>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.CustomXMLPart>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.CustomXMLPart>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.CustomXMLPart>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.CustomXMLPart> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.CustomXMLPart item in innerEnumerator)
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