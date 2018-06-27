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
	/// DispatchInterface Controls 
	/// SupportByVersion MSForms, 2
	/// </summary>
	public class Controls : COMObject, NetOffice.MSFormsApi.Controls
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
                    _contractType = typeof(NetOffice.MSFormsApi.Controls);
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
                    _type = typeof(Controls);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Controls() : base()
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
		public virtual void Clear()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="cx">Int32 cx</param>
		/// <param name="cy">Int32 cy</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void _Move(Int32 cx, Int32 cy)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_Move", cx, cy);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public virtual void SelectAll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SelectAll");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clsid">Int32 clsid</param>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Control _AddByClass(Int32 clsid)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "_AddByClass", typeof(NetOffice.MSFormsApi.Control), clsid);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public virtual void AlignToGrid()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AlignToGrid");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public virtual void BringForward()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BringForward");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public virtual void BringToFront()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BringToFront");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public virtual void Copy()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public virtual void Cut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Cut");
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
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Control _GetItemByIndex(Int32 lIndex)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "_GetItemByIndex", typeof(NetOffice.MSFormsApi.Control), lIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="pstr">string pstr</param>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Control _GetItemByName(string pstr)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "_GetItemByName", typeof(NetOffice.MSFormsApi.Control), pstr);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="iD">Int32 iD</param>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Control _GetItemByID(Int32 iD)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "_GetItemByID", typeof(NetOffice.MSFormsApi.Control), iD);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void SendBackward()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendBackward");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void SendToBack()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendToBack");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="cx">Single cx</param>
		/// <param name="cy">Single cy</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void Move(Single cx, Single cy)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", cx, cy);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrProgID">string bstrProgID</param>
		/// <param name="name">optional object name</param>
		/// <param name="visible">optional object visible</param>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Control Add(string bstrProgID, object name, object visible)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "Add", typeof(NetOffice.MSFormsApi.Control), bstrProgID, name, visible);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrProgID">string bstrProgID</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Control Add(string bstrProgID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "Add", typeof(NetOffice.MSFormsApi.Control), bstrProgID);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrProgID">string bstrProgID</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Control Add(string bstrProgID, object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "Add", typeof(NetOffice.MSFormsApi.Control), bstrProgID, name);
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

