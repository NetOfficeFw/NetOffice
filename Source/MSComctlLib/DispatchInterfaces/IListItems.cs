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
	/// DispatchInterface IListItems 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class IListItems : COMObject, IEnumerableProvider<NetOffice.MSComctlLibApi.IListItem>
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
                    _type = typeof(IListItems);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IListItems(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IListItems(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListItems(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListItems(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListItems(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListItems(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListItems() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListItems(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Count", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSComctlLibApi.IListItem get_ControlDefault(object index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSComctlLibApi.IListItem>(this, "ControlDefault", NetOffice.MSComctlLibApi.IListItem.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6), Redirect("get_ControlDefault")]
		public NetOffice.MSComctlLibApi.IListItem ControlDefault(object index)
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
		public NetOffice.MSComctlLibApi.IListItem this[object index]
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.IListItem>(this, "Item", index);
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
		/// <param name="icon">optional object icon</param>
		/// <param name="smallIcon">optional object smallIcon</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		public NetOffice.MSComctlLibApi.IListItem Add(object index, object key, object text, object icon, object smallIcon)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "Add", new object[]{ index, key, text, icon, smallIcon });
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListItem Add()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "Add");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListItem Add(object index)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "Add", index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListItem Add(object index, object key)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "Add", index, key);
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
		public NetOffice.MSComctlLibApi.IListItem Add(object index, object key, object text)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "Add", index, key, text);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="icon">optional object icon</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListItem Add(object index, object key, object text, object icon)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "Add", index, key, text, icon);
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
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		public void Remove(object index)
		{
			 Factory.ExecuteMethod(this, "Remove", index);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.MSComctlLibApi.IListItem>

        ICOMObject IEnumerableProvider<NetOffice.MSComctlLibApi.IListItem>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.MSComctlLibApi.IListItem>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.MSComctlLibApi.IListItem>

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public IEnumerator<NetOffice.MSComctlLibApi.IListItem> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.MSComctlLibApi.IListItem item in innerEnumerator)
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
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion

		#pragma warning restore
	}
}