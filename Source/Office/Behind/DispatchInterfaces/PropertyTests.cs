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
    /// DispatchInterface PropertyTests 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
    public class PropertyTests : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.PropertyTests
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
                    _type = typeof(PropertyTests);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PropertyTests() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public NetOffice.OfficeApi.PropertyTest this[Int32 index]
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PropertyTest>(this, "Item", typeof(NetOffice.OfficeApi.PropertyTest), index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Count
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="condition">NetOffice.OfficeApi.Enums.MsoCondition condition</param>
        /// <param name="value">optional object value</param>
        /// <param name="secondValue">optional object secondValue</param>
        /// <param name="connector">optional NetOffice.OfficeApi.Enums.MsoConnector Connector = 1</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Add(string name, NetOffice.OfficeApi.Enums.MsoCondition condition, object value, object secondValue, object connector)
        {
            Factory.ExecuteMethod(this, "Add", new object[] { name, condition, value, secondValue, connector });
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="condition">NetOffice.OfficeApi.Enums.MsoCondition condition</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Add(string name, NetOffice.OfficeApi.Enums.MsoCondition condition)
        {
            Factory.ExecuteMethod(this, "Add", name, condition);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="condition">NetOffice.OfficeApi.Enums.MsoCondition condition</param>
        /// <param name="value">optional object value</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Add(string name, NetOffice.OfficeApi.Enums.MsoCondition condition, object value)
        {
            Factory.ExecuteMethod(this, "Add", name, condition, value);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="condition">NetOffice.OfficeApi.Enums.MsoCondition condition</param>
        /// <param name="value">optional object value</param>
        /// <param name="secondValue">optional object secondValue</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Add(string name, NetOffice.OfficeApi.Enums.MsoCondition condition, object value, object secondValue)
        {
            Factory.ExecuteMethod(this, "Add", name, condition, value, secondValue);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Remove(Int32 index)
        {
            Factory.ExecuteMethod(this, "Remove", index);
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.PropertyTest>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.PropertyTest>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.PropertyTest>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.PropertyTest>

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.PropertyTest> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.PropertyTest item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
