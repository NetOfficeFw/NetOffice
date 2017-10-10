using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSComctlLibApi
{
	/// <summary>
	/// DispatchInterface ITabs 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class ITabs : COMObject, IEnumerableProvider<NetOffice.MSComctlLibApi.ITab>
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
                    _type = typeof(ITabs);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ITabs(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ITabs(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabs(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabs(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabs(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabs(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabs() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabs(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public Int16 Count
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Count");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Count", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSComctlLibApi.ITab get_ControlDefault(object pvIndex)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSComctlLibApi.ITab>(this, "ControlDefault", NetOffice.MSComctlLibApi.ITab.LateBindingApiWrapperType, pvIndex);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ControlDefault(object pvIndex, NetOffice.MSComctlLibApi.ITab value)
		{
			Factory.ExecutePropertySet(this, "ControlDefault", pvIndex, value);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersion("MSComctlLib", 6), Redirect("get_ControlDefault")]
		public NetOffice.MSComctlLibApi.ITab ControlDefault(object pvIndex)
		{
			return get_ControlDefault(pvIndex);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.MSComctlLibApi.ITab this[object pvIndex]
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.ITab>(this, "Item", pvIndex);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Item", value, pvIndex);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersion("MSComctlLib", 6)]
		public void Remove(object pvIndex)
		{
			 Factory.ExecuteMethod(this, "Remove", pvIndex);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public void Clear()
		{
			 Factory.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="pvIndex">optional object pvIndex</param>
		/// <param name="pvKey">optional object pvKey</param>
		/// <param name="pvCaption">optional object pvCaption</param>
		/// <param name="pvImage">optional object pvImage</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		public NetOffice.MSComctlLibApi.ITab Add(object pvIndex, object pvKey, object pvCaption, object pvImage)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.ITab>(this, "Add", pvIndex, pvKey, pvCaption, pvImage);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.ITab Add()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.ITab>(this, "Add");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="pvIndex">optional object pvIndex</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.ITab Add(object pvIndex)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.ITab>(this, "Add", pvIndex);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="pvIndex">optional object pvIndex</param>
		/// <param name="pvKey">optional object pvKey</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.ITab Add(object pvIndex, object pvKey)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.ITab>(this, "Add", pvIndex, pvKey);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="pvIndex">optional object pvIndex</param>
		/// <param name="pvKey">optional object pvKey</param>
		/// <param name="pvCaption">optional object pvCaption</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.ITab Add(object pvIndex, object pvKey, object pvCaption)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.ITab>(this, "Add", pvIndex, pvKey, pvCaption);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.MSComctlLibApi.ITab>

        ICOMObject IEnumerableProvider<NetOffice.MSComctlLibApi.ITab>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.MSComctlLibApi.ITab>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.MSComctlLibApi.ITab>

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public IEnumerator<NetOffice.MSComctlLibApi.ITab> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.MSComctlLibApi.ITab item in innerEnumerator)
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