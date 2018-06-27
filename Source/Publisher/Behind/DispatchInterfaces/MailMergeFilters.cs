using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface MailMergeFilters 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	public class MailMergeFilters : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.PublisherApi.MailMergeFilters
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
                    _contractType = typeof(NetOffice.PublisherApi.MailMergeFilters);
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
                    _type = typeof(MailMergeFilters);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public MailMergeFilters() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual object this[Int32 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Item", index);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="column">string column</param>
		/// <param name="comparison">NetOffice.OfficeApi.Enums.MsoFilterComparison comparison</param>
		/// <param name="conjunction">NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction</param>
		/// <param name="bstrCompareTo">optional string bstrCompareTo = </param>
		/// <param name="deferUpdate">optional bool DeferUpdate = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Add(string column, NetOffice.OfficeApi.Enums.MsoFilterComparison comparison, NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction, object bstrCompareTo, object deferUpdate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", new object[]{ column, comparison, conjunction, bstrCompareTo, deferUpdate });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="column">string column</param>
		/// <param name="comparison">NetOffice.OfficeApi.Enums.MsoFilterComparison comparison</param>
		/// <param name="conjunction">NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Add(string column, NetOffice.OfficeApi.Enums.MsoFilterComparison comparison, NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", column, comparison, conjunction);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="column">string column</param>
		/// <param name="comparison">NetOffice.OfficeApi.Enums.MsoFilterComparison comparison</param>
		/// <param name="conjunction">NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction</param>
		/// <param name="bstrCompareTo">optional string bstrCompareTo = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Add(string column, NetOffice.OfficeApi.Enums.MsoFilterComparison comparison, NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction, object bstrCompareTo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Add", column, comparison, conjunction, bstrCompareTo);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="deferUpdate">optional bool DeferUpdate = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Delete(Int32 index, object deferUpdate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete", index, deferUpdate);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Delete(Int32 index)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete", index);
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
        /// SupportByVersion Publisher, 14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [CustomEnumerator]
        public virtual IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
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
