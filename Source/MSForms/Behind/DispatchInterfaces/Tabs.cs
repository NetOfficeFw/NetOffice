using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.MSFormsApi;

namespace NetOffice.MSFormsApi.Behind
{
	/// <summary>
	/// DispatchInterface Tabs 
	/// SupportByVersion MSForms, 2
	/// </summary>
	public class Tabs : COMObject, NetOffice.MSFormsApi.Tabs
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
                    _contractType = typeof(NetOffice.MSFormsApi.Tabs);
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
                    _type = typeof(Tabs);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Tabs() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Tab _GetItemByIndex(Int32 lIndex)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Tab>(this, "_GetItemByIndex", typeof(NetOffice.MSFormsApi.Tab), lIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstr">string bstr</param>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Tab _GetItemByName(string bstr)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Tab>(this, "_GetItemByName", typeof(NetOffice.MSFormsApi.Tab), bstr);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="varg">object varg</param>
		[SupportByVersion("MSForms", 2)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual object this[object varg]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Item", varg);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual object Enum()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Enum");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="lIndex">optional object lIndex</param>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Tab Add(object bstrName, object bstrCaption, object lIndex)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Tab>(this, "Add", typeof(NetOffice.MSFormsApi.Tab), bstrName, bstrCaption, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Tab Add()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Tab>(this, "Add", typeof(NetOffice.MSFormsApi.Tab));
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Tab Add(object bstrName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Tab>(this, "Add", typeof(NetOffice.MSFormsApi.Tab), bstrName);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Tab Add(object bstrName, object bstrCaption)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Tab>(this, "Add", typeof(NetOffice.MSFormsApi.Tab), bstrName, bstrCaption);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Tab _Add(string bstrName, string bstrCaption)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Tab>(this, "_Add", typeof(NetOffice.MSFormsApi.Tab), bstrName, bstrCaption);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Tab _Insert(string bstrName, string bstrCaption, Int32 lIndex)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Tab>(this, "_Insert", typeof(NetOffice.MSFormsApi.Tab), bstrName, bstrCaption, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="varg">object varg</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void Remove(object varg)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Remove", varg);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void Clear()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Clear");
		}

        #endregion

        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<object>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);
        }

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion MSForms, 2
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        public virtual IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion MSForms, 2
        /// </summary>
        [SupportByVersion("MSForms", 2)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, true);
		}

		#endregion

		#pragma warning restore
	}
}

