using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface Pages 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Pages : COMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<object>
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
                    _type = typeof(Pages);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Pages(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Pages(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(string progId) : base(progId)
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
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="varg">object varg</param>
		[SupportByVersion("MSForms", 2)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public object this[object varg]
		{
			get
			{
				return Factory.ExecuteVariantMethodGet(this, "Item", varg);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public object Enum()
		{
			return Factory.ExecuteVariantMethodGet(this, "Enum");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="lIndex">optional object lIndex</param>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Page Add(object bstrName, object bstrCaption, object lIndex)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Page>(this, "Add", NetOffice.MSFormsApi.Page.LateBindingApiWrapperType, bstrName, bstrCaption, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Page Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Page>(this, "Add", NetOffice.MSFormsApi.Page.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Page Add(object bstrName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Page>(this, "Add", NetOffice.MSFormsApi.Page.LateBindingApiWrapperType, bstrName);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Page Add(object bstrName, object bstrCaption)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Page>(this, "Add", NetOffice.MSFormsApi.Page.LateBindingApiWrapperType, bstrName, bstrCaption);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clsid">Int32 clsid</param>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Page _AddCtrl(Int32 clsid, string bstrName, string bstrCaption)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Page>(this, "_AddCtrl", NetOffice.MSFormsApi.Page.LateBindingApiWrapperType, clsid, bstrName, bstrCaption);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clsid">Int32 clsid</param>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Page _InsertCtrl(Int32 clsid, string bstrName, string bstrCaption, Int32 lIndex)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Page>(this, "_InsertCtrl", NetOffice.MSFormsApi.Page.LateBindingApiWrapperType, clsid, bstrName, bstrCaption, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Control _GetItemByIndex(Int32 lIndex)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "_GetItemByIndex", NetOffice.MSFormsApi.Control.LateBindingApiWrapperType, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="pstrName">string pstrName</param>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Control _GetItemByName(string pstrName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "_GetItemByName", NetOffice.MSFormsApi.Control.LateBindingApiWrapperType, pstrName);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="varg">object varg</param>
		[SupportByVersion("MSForms", 2)]
		public void Remove(object varg)
		{
			 Factory.ExecuteMethod(this, "Remove", varg);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void Clear()
		{
			 Factory.ExecuteMethod(this, "Clear");
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
        public IEnumerator<object> GetEnumerator()
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