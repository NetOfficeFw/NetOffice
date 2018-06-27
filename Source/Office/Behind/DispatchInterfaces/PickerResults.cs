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
    /// DispatchInterface PickerResults 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864136.aspx </remarks>
    public class PickerResults : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.PickerResults
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
                    _contractType = typeof(NetOffice.OfficeApi.PickerResults);
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
                    _type = typeof(PickerResults);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PickerResults() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.OfficeApi.PickerResult this[Int32 index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PickerResult>(this, "Item", typeof(NetOffice.OfficeApi.PickerResult), index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865190.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864663.aspx </remarks>
        /// <param name="id">string id</param>
        /// <param name="displayName">string displayName</param>
        /// <param name="type">string type</param>
        /// <param name="sIPId">optional string SIPId = </param>
        /// <param name="itemData">optional object itemData</param>
        /// <param name="subItems">optional object subItems</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PickerResult Add(string id, string displayName, string type, object sIPId, object itemData, object subItems)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResult>(this, "Add", typeof(NetOffice.OfficeApi.PickerResult), new object[] { id, displayName, type, sIPId, itemData, subItems });
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864663.aspx </remarks>
        /// <param name="id">string id</param>
        /// <param name="displayName">string displayName</param>
        /// <param name="type">string type</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PickerResult Add(string id, string displayName, string type)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResult>(this, "Add", typeof(NetOffice.OfficeApi.PickerResult), id, displayName, type);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864663.aspx </remarks>
        /// <param name="id">string id</param>
        /// <param name="displayName">string displayName</param>
        /// <param name="type">string type</param>
        /// <param name="sIPId">optional string SIPId = </param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PickerResult Add(string id, string displayName, string type, object sIPId)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResult>(this, "Add", typeof(NetOffice.OfficeApi.PickerResult), id, displayName, type, sIPId);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864663.aspx </remarks>
        /// <param name="id">string id</param>
        /// <param name="displayName">string displayName</param>
        /// <param name="type">string type</param>
        /// <param name="sIPId">optional string SIPId = </param>
        /// <param name="itemData">optional object itemData</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PickerResult Add(string id, string displayName, string type, object sIPId, object itemData)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResult>(this, "Add", typeof(NetOffice.OfficeApi.PickerResult), new object[] { id, displayName, type, sIPId, itemData });
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.PickerResult>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.PickerResult>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.PickerResult>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.PickerResult>

        /// <summary>
        /// SupportByVersion Office, 14,15,16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual IEnumerator<NetOffice.OfficeApi.PickerResult> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.PickerResult item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 14,15,16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
