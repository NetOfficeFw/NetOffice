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
	/// DispatchInterface Controls 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Controls : COMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<object>
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
                    _type = typeof(Controls);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Controls(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Controls(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Controls(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Controls(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Controls(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Controls(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Controls() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Controls(string progId) : base(progId)
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
		public void Clear()
		{
			 Factory.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="cx">Int32 cx</param>
		/// <param name="cy">Int32 cy</param>
		[SupportByVersion("MSForms", 2)]
		public void _Move(Int32 cx, Int32 cy)
		{
			 Factory.ExecuteMethod(this, "_Move", cx, cy);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public void SelectAll()
		{
			 Factory.ExecuteMethod(this, "SelectAll");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clsid">Int32 clsid</param>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Control _AddByClass(Int32 clsid)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "_AddByClass", NetOffice.MSFormsApi.Control.LateBindingApiWrapperType, clsid);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public void AlignToGrid()
		{
			 Factory.ExecuteMethod(this, "AlignToGrid");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public void BringForward()
		{
			 Factory.ExecuteMethod(this, "BringForward");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public void BringToFront()
		{
			 Factory.ExecuteMethod(this, "BringToFront");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
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
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Control _GetItemByIndex(Int32 lIndex)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "_GetItemByIndex", NetOffice.MSFormsApi.Control.LateBindingApiWrapperType, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="pstr">string pstr</param>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Control _GetItemByName(string pstr)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "_GetItemByName", NetOffice.MSFormsApi.Control.LateBindingApiWrapperType, pstr);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="iD">Int32 iD</param>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Control _GetItemByID(Int32 iD)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "_GetItemByID", NetOffice.MSFormsApi.Control.LateBindingApiWrapperType, iD);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void SendBackward()
		{
			 Factory.ExecuteMethod(this, "SendBackward");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void SendToBack()
		{
			 Factory.ExecuteMethod(this, "SendToBack");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="cx">Single cx</param>
		/// <param name="cy">Single cy</param>
		[SupportByVersion("MSForms", 2)]
		public void Move(Single cx, Single cy)
		{
			 Factory.ExecuteMethod(this, "Move", cx, cy);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrProgID">string bstrProgID</param>
		/// <param name="name">optional object name</param>
		/// <param name="visible">optional object visible</param>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Control Add(string bstrProgID, object name, object visible)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "Add", NetOffice.MSFormsApi.Control.LateBindingApiWrapperType, bstrProgID, name, visible);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrProgID">string bstrProgID</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Control Add(string bstrProgID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "Add", NetOffice.MSFormsApi.Control.LateBindingApiWrapperType, bstrProgID);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrProgID">string bstrProgID</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Control Add(string bstrProgID, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSFormsApi.Control>(this, "Add", NetOffice.MSFormsApi.Control.LateBindingApiWrapperType, bstrProgID, name);
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