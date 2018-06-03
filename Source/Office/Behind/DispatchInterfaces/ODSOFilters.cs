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
    /// DispatchInterface ODSOFilters 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865224.aspx </remarks>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Variant, EnumeratorInvoke.Custom), HasIndexProperty(IndexInvoke.Method, "Item")]
    public class ODSOFilters : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.ODSOFilters
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
                    _type = typeof(ODSOFilters);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ODSOFilters() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860835.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public Int32 Count
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861525.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public object this[Int32 index]
        {
            get
            {
                return Factory.ExecuteVariantMethodGet(this, "Item", index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864658.aspx </remarks>
        /// <param name="column">string column</param>
        /// <param name="comparison">NetOffice.OfficeApi.Enums.MsoFilterComparison comparison</param>
        /// <param name="conjunction">NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction</param>
        /// <param name="bstrCompareTo">optional string bstrCompareTo = </param>
        /// <param name="deferUpdate">optional bool DeferUpdate = false</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public void Add(string column, NetOffice.OfficeApi.Enums.MsoFilterComparison comparison, NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction, object bstrCompareTo, object deferUpdate)
        {
            Factory.ExecuteMethod(this, "Add", new object[] { column, comparison, conjunction, bstrCompareTo, deferUpdate });
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864658.aspx </remarks>
        /// <param name="column">string column</param>
        /// <param name="comparison">NetOffice.OfficeApi.Enums.MsoFilterComparison comparison</param>
        /// <param name="conjunction">NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public void Add(string column, NetOffice.OfficeApi.Enums.MsoFilterComparison comparison, NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction)
        {
            Factory.ExecuteMethod(this, "Add", column, comparison, conjunction);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864658.aspx </remarks>
        /// <param name="column">string column</param>
        /// <param name="comparison">NetOffice.OfficeApi.Enums.MsoFilterComparison comparison</param>
        /// <param name="conjunction">NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction</param>
        /// <param name="bstrCompareTo">optional string bstrCompareTo = </param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public void Add(string column, NetOffice.OfficeApi.Enums.MsoFilterComparison comparison, NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction, object bstrCompareTo)
        {
            Factory.ExecuteMethod(this, "Add", column, comparison, conjunction, bstrCompareTo);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860318.aspx </remarks>
        /// <param name="index">Int32 index</param>
        /// <param name="deferUpdate">optional bool DeferUpdate = false</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public void Delete(Int32 index, object deferUpdate)
        {
            Factory.ExecuteMethod(this, "Delete", index, deferUpdate);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860318.aspx </remarks>
        /// <param name="index">Int32 index</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public void Delete(Int32 index)
        {
            Factory.ExecuteMethod(this, "Delete", index);
        }

        #endregion

        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return this;
        }

        IEnumerable IEnumerableProvider<object>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion Office, 10,11,12,14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        [CustomEnumerator]
        public IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 10,11,12,14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        [CustomEnumerator]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            int count = Count;
            object[] enumeratorObjects = new object[count];
            for (int i = 0; i < count; i++)
                enumeratorObjects[i] = this[i + 1];

            foreach (object item in enumeratorObjects)
                yield return item;
        }

        #endregion

        #pragma warning restore
    }
}
