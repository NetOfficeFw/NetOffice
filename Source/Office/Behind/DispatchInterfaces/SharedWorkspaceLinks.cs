using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface SharedWorkspaceLinks 
    /// SupportByVersion Office, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863849.aspx </remarks>
    [SupportByVersion("Office", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
    public class SharedWorkspaceLinks : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.SharedWorkspaceLinks
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
                    _type = typeof(SharedWorkspaceLinks);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public SharedWorkspaceLinks() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public NetOffice.OfficeApi.SharedWorkspaceLink this[Int32 index]
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspaceLink>(this, "Item", typeof(NetOffice.OfficeApi.SharedWorkspaceLink), index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864175.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public Int32 Count
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861528.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861770.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public bool ItemCountExceeded
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ItemCountExceeded");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862533.aspx </remarks>
        /// <param name="uRL">string uRL</param>
        /// <param name="description">optional object description</param>
        /// <param name="notes">optional object notes</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.SharedWorkspaceLink Add(string uRL, object description, object notes)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceLink>(this, "Add", typeof(NetOffice.OfficeApi.SharedWorkspaceLink), uRL, description, notes);
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862533.aspx </remarks>
        /// <param name="uRL">string uRL</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.SharedWorkspaceLink Add(string uRL)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceLink>(this, "Add", typeof(NetOffice.OfficeApi.SharedWorkspaceLink), uRL);
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862533.aspx </remarks>
        /// <param name="uRL">string uRL</param>
        /// <param name="description">optional object description</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.SharedWorkspaceLink Add(string uRL, object description)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceLink>(this, "Add", typeof(NetOffice.OfficeApi.SharedWorkspaceLink), uRL, description);
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceLink>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceLink>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceLink>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.SharedWorkspaceLink>

        /// <summary>
        /// SupportByVersion Office, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.SharedWorkspaceLink> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.SharedWorkspaceLink item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
