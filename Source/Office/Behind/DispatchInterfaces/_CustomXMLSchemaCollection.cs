using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface _CustomXMLSchemaCollection 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    public class _CustomXMLSchemaCollection : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi._CustomXMLSchemaCollection
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OfficeApi._CustomXMLSchemaCollection);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _CustomXMLSchemaCollection() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862208.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864015.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.OfficeApi.CustomXMLSchema this[object index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLSchema>(this, "Item", typeof(NetOffice.OfficeApi.CustomXMLSchema), index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860878.aspx </remarks>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_NamespaceURI(Int32 index)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NamespaceURI", index);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_NamespaceURI
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860878.aspx </remarks>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_NamespaceURI")]
        public virtual string NamespaceURI(Int32 index)
        {
            return get_NamespaceURI(index);
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864881.aspx </remarks>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="alias">optional string Alias = </param>
        /// <param name="fileName">optional string FileName = </param>
        /// <param name="installForAllUsers">optional bool InstallForAllUsers = false</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLSchema Add(object namespaceURI, object alias, object fileName, object installForAllUsers)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLSchema>(this, "Add", typeof(NetOffice.OfficeApi.CustomXMLSchema), namespaceURI, alias, fileName, installForAllUsers);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864881.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLSchema Add()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLSchema>(this, "Add", typeof(NetOffice.OfficeApi.CustomXMLSchema));
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864881.aspx </remarks>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLSchema Add(object namespaceURI)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLSchema>(this, "Add", typeof(NetOffice.OfficeApi.CustomXMLSchema), namespaceURI);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864881.aspx </remarks>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="alias">optional string Alias = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLSchema Add(object namespaceURI, object alias)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLSchema>(this, "Add", typeof(NetOffice.OfficeApi.CustomXMLSchema), namespaceURI, alias);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864881.aspx </remarks>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="alias">optional string Alias = </param>
        /// <param name="fileName">optional string FileName = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLSchema Add(object namespaceURI, object alias, object fileName)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLSchema>(this, "Add", typeof(NetOffice.OfficeApi.CustomXMLSchema), namespaceURI, alias, fileName);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864690.aspx </remarks>
        /// <param name="schemaCollection">NetOffice.OfficeApi.CustomXMLSchemaCollection schemaCollection</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AddCollection(NetOffice.OfficeApi.CustomXMLSchemaCollection schemaCollection)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddCollection", schemaCollection);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864142.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool Validate()
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Validate");
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.CustomXMLSchema>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.CustomXMLSchema>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
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
        public virtual IEnumerator<NetOffice.OfficeApi.CustomXMLSchema> GetEnumerator()
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
