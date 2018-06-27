using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.MSComctlLibApi;

namespace NetOffice.MSComctlLibApi.Behind
{
	/// <summary>
	/// DispatchInterface IColumnHeaders 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	public class IColumnHeaders : COMObject, NetOffice.MSComctlLibApi.IColumnHeaders
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
                    _contractType = typeof(NetOffice.MSComctlLibApi.IColumnHeaders);
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
                    _type = typeof(IColumnHeaders);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IColumnHeaders() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Count", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader get_ControlDefault(object index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "ControlDefault", typeof(NetOffice.MSComctlLibApi.IColumnHeader), index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6), Redirect("get_ControlDefault")]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader ControlDefault(object index)
		{
			return get_ControlDefault(index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Item", index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="width">optional object width</param>
		/// <param name="alignment">optional object alignment</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index, object key, object text, object width, object alignment)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add_PreVB98", new object[]{ index, key, text, width, alignment });
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add_PreVB98");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add_PreVB98", index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index, object key)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add_PreVB98", index, key);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index, object key, object text)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add_PreVB98", index, key, text);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="width">optional object width</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index, object key, object text, object width)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add_PreVB98", index, key, text, width);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void Clear()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void Remove(object index)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Remove", index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="width">optional object width</param>
		/// <param name="alignment">optional object alignment</param>
		/// <param name="icon">optional object icon</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key, object text, object width, object alignment, object icon)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add", new object[]{ index, key, text, width, alignment, icon });
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add(object index)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add", index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add", index, key);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key, object text)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add", index, key, text);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key, object text, object width)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add", index, key, text, width);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="width">optional object width</param>
		/// <param name="alignment">optional object alignment</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key, object text, object width, object alignment)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IColumnHeader>(this, "Add", new object[]{ index, key, text, width, alignment });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.MSComctlLibApi.IColumnHeader>

        ICOMObject IEnumerableProvider<NetOffice.MSComctlLibApi.IColumnHeader>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.MSComctlLibApi.IColumnHeader>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.MSComctlLibApi.IColumnHeader>

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public virtual IEnumerator<NetOffice.MSComctlLibApi.IColumnHeader> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.MSComctlLibApi.IColumnHeader item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

